'use client';

import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  Database, 
  FileText, 
  UploadCloud, 
  BarChart2, 
  MousePointerClick, 
  ListFilter, 
  Calendar, 
  MapPin, 
  Search, 
  Copy, 
  EyeOff, 
  Download, 
  CheckCircle2, 
  RefreshCw,
  Printer 
} from 'lucide-react';

// 초기 샘플 데이터 (파일 업로드 전 즉시 동작 및 시각화 확인용)
const INITIAL_SAMPLE_DATA = [
  { '구분': '정기보수', '장치명': '궤도회로장치', '개소': '서울역 구내', '세부작업내용': '궤도회로 송수신기 전압 레벨 측정 및 노이즈 필터 접점 클리닝', '발생일': '2026-03-01' },
  { '구분': '긴급점검', '장치명': '선로전환기', '개소': '용산역 상행선', '세부작업내용': '선로전환기 밀착 검지 센서 오차 교정 및 쇄정상태 재조정', '발생일': '2026-03-02' },
  { '구분': '정기보수', '장치명': '신호발신기', '개소': '광명역 진입로', '세부작업내용': 'LED 신호등 렌즈 먼지 제거 및 램프 전류치 정기 점검', '발생일': '2026-03-03' },
  { '구분': '장애조치', '장치명': '궤도회로장치', '개소': '서울역 구내', '세부작업내용': '궤도회로 송수신기 전압 레벨 측정 및 노이즈 필터 접점 클리닝', '발생일': '2026-03-04' },
  { '구분': '정기보수', '장치명': '자동폐색장치', '개소': '천안아산역 구간', '세부작업내용': '폐색 계전기 반응속도 시험 및 전원선 절연저항 측정', '발생일': '2026-03-05' },
  { '구분': '부품교체', '장치명': '선로전환기', '개소': '대전역 분기부', '세부작업내용': '전환 모터 브러시 마모 확인 후 신품 교체 및 급유 작업', '발생일': '2026-03-06' },
  { '구분': '긴급점검', '장치명': '궤도회로장치', '개소': '동대구역 구내', '세부작업내용': '임피던스 본드 연결선 이완 체결 및 접지저항 정밀 검측', '발생일': '2026-03-07' },
  { '구분': '정기보수', '장치명': '열차제어장치(ATP)', '개소': '신경주역 하행선', '세부작업내용': '지상 발리스(Balise) 신호 송출 상태 시험 및 결선 상태 점검', '발생일': '2026-03-08' },
  { '구분': '장애조치', '장치명': '선로전환기', '개소': '용산역 상행선', '세부작업내용': '선로전환기 밀착 검지 센서 오차 교정 및 쇄정상태 재조정', '발생일': '2026-03-09' },
  { '구분': '정기보수', '장치명': '신호발신기', '개소': '부산역 구내', '세부작업내용': '중계 신호기 점등 시험 및 제어반 릴레이 동작 검사', '발생일': '2026-03-10' },
  { '구분': '정기보수', '장치명': '궤도회로장치', '개소': '서울역 구내', '세부작업내용': '궤도회로 송수신기 전압 레벨 측정 및 노이즈 필터 접점 클리닝', '발생일': '2026-03-11' }
];

export default function MaintenanceAnalysisDashboard() {
  const [tableData, setTableData] = useState<any[]>(INITIAL_SAMPLE_DATA);
  const [columns, setColumns] = useState<string[]>(['구분', '장치명', '개소', '세부작업내용', '발생일']);
  const [fileName, setFileName] = useState<string>("유지보수실적_샘플데이터.xlsx");
  const [summary, setSummary] = useState<string>('');
  const [isSummarizing, setIsSummarizing] = useState<boolean>(false);
  const [showSummary, setShowSummary] = useState<boolean>(false);
  const [copied, setCopied] = useState<boolean>(false);
  
  // 필터 상태
  const [selectedDevice, setSelectedDevice] = useState<string>("전체");
  const [selectedLocation, setSelectedLocation] = useState<string>("전체");
  const [searchText, setSearchText] = useState<string>("");

  // 장치 선택 처리 핸들러 (장치 변경 시 연관 개소 필터 즉시 "전체"로 초기화)
  const handleDeviceSelect = (device: string) => {
    setSelectedDevice(device);
    setSelectedLocation("전체");
  };

  // 1. 파일 업로드 및 인코딩 방어 파싱 로직
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const file = files[0];
    setFileName(file.name);
    setSelectedDevice("전체");
    setSelectedLocation("전체");
    setSearchText("");
    setSummary('');
    setShowSummary(false);
    
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const buffer = event.target?.result as ArrayBuffer;
        if (!buffer || buffer.byteLength === 0) {
          alert('파일 내용이 비어 있습니다.');
          return;
        }

        let workbook: XLSX.WorkBook | null = null;
        const isCsv = file.name.toLowerCase().endsWith('.csv') || file.name.toLowerCase().endsWith('.txt');

        if (isCsv) {
          // CSV 파일: UTF-8 실패 시 EUC-KR 이중 디코딩 방어
          let text = '';
          try {
            const utf8Decoder = new TextDecoder('utf-8', { fatal: true });
            text = utf8Decoder.decode(buffer);
          } catch {
            const eucKrDecoder = new TextDecoder('euc-kr');
            text = eucKrDecoder.decode(buffer);
          }
          workbook = XLSX.read(text, { type: 'string' });
        } else {
          // .xlsx / .xls 등 바이너리 엑셀 파일: Uint8Array 기반 파싱
          try {
            const data = new Uint8Array(buffer);
            workbook = XLSX.read(data, { type: 'array', cellDates: true, codepage: 949 });
          } catch (binErr) {
            console.warn('Binary read failed, trying text decoder fallback:', binErr);
            // 텍스트(HTML/CSV)가 .xls 확장자로 저장된 경우를 위한 폴백
            let text = '';
            try {
              const utf8Decoder = new TextDecoder('utf-8', { fatal: true });
              text = utf8Decoder.decode(buffer);
            } catch {
              const eucKrDecoder = new TextDecoder('euc-kr');
              text = eucKrDecoder.decode(buffer);
            }
            workbook = XLSX.read(text, { type: 'string' });
          }
        }

        if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
          throw new Error('시트 정보를 읽을 수 없습니다.');
        }

        // 첫 번째 유효 시트 선택
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        if (!worksheet) {
          throw new Error(`시트 [${firstSheetName}] 데이터를 찾을 수 없습니다.`);
        }

        // 2D 배열로 변환
        const rows: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        if (!rows || rows.length === 0) {
          alert('시트에 유효한 데이터가 없습니다.');
          return;
        }

        // 헤더 행 자동 탐색 (상위 15개 행 중 키워드 포함 행 탐색)
        let headerIdx = 0;
        for (let i = 0; i < Math.min(rows.length, 15); i++) {
          const rowStr = rows[i].map(c => String(c || '').trim()).join(',');
          if (
            rowStr.includes('구분') || 
            rowStr.includes('장치명') || 
            rowStr.includes('장치') || 
            rowStr.includes('개소') || 
            rowStr.includes('작업') || 
            rowStr.includes('발생일') ||
            rowStr.includes('일자')
          ) {
            headerIdx = i;
            break;
          }
        }

        const rawHeaders = rows[headerIdx] || [];
        const validCols: { index: number; name: string }[] = [];
        const excludeCols = ['조치여부', '조치일', '소요물품', 'A/S', '조치내용', '처리상태'];

        rawHeaders.forEach((h, idx) => {
          const colName = h ? String(h).trim() : '';
          if (colName && !colName.startsWith('__EMPTY') && !excludeCols.includes(colName)) {
            validCols.push({ index: idx, name: colName });
          }
        });

        // 만약 헤더가 제대로 검출되지 않았을 경우 fallback 컬럼명 생성
        if (validCols.length === 0) {
          const maxCols = Math.max(...rows.slice(headerIdx).map(r => r.length), 0);
          for (let c = 0; c < maxCols; c++) {
            validCols.push({ index: c, name: `열_${c + 1}` });
          }
        }

        const dataRows = rows.slice(headerIdx + 1).filter(row => 
          row && row.some(cell => cell !== undefined && cell !== null && String(cell).trim() !== '')
        );

        const formatted = dataRows.map(row => {
          const obj: Record<string, any> = {};
          validCols.forEach(col => {
            let val = row[col.index];
            if (val instanceof Date) {
              val = val.toISOString().slice(0, 10);
            }
            obj[col.name] = val !== undefined && val !== null ? String(val).trim() : '';
          });
          return obj;
        });

        setColumns(validCols.map(c => c.name));
        setTableData(formatted);
      } catch (err: any) {
        console.error('File Parse Error:', err);
        alert(`파일을 분석하는 중 오류가 발생했습니다: ${err.message || '파일 형식을 확인해 주세요.'}`);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // 핵심 컬럼 키 탐색 (B열/장치명, 개소, 발생일, 세부작업내용)
  const deviceKey = useMemo(() => {
    if (columns.length === 0) return '';
    const found = columns.find(c => c.includes('장치명') || c.includes('장치') || c.includes('기기') || c.includes('설비'));
    if (found) return found;
    return columns.length > 1 ? columns[1] : columns[0]; // B열(두번째 열) 기준
  }, [columns]);

  const locationKey = useMemo(() => {
    if (columns.length === 0) return '';
    return columns.find(c => c.includes('개소') || c.includes('지역') || c.includes('위치') || c.includes('역명') || c.includes('역')) || '';
  }, [columns]);

  const dateKey = useMemo(() => {
    if (columns.length === 0) return '';
    return columns.find(c => c.includes('발생일') || c.includes('일자') || c.includes('날짜') || c.includes('등록일') || c.includes('일시')) || '';
  }, [columns]);

  const workContentKey = useMemo(() => {
    if (columns.length === 0) return '';
    return columns.find(c => c.includes('세부작업') || c.includes('작업내용') || c.includes('세부내용') || c.includes('내용') || c.includes('현상')) || '';
  }, [columns]);

  // 2. 장치명 가나다순 목록
  const deviceList = useMemo(() => {
    if (tableData.length === 0 || !deviceKey) return [];
    const devices = new Set<string>();
    tableData.forEach(row => {
      const val = row[deviceKey];
      if (val !== undefined && val !== null) {
        const cleanVal = String(val).trim();
        if (cleanVal !== '' && cleanVal.length < 20) devices.add(cleanVal);
      }
    });
    return Array.from(devices).sort((a, b) => a.localeCompare(b, 'ko'));
  }, [tableData, deviceKey]);

  // 개소 목록 (선택된 장치가 있으면 해당 장치에 속한 개소들만 동적으로 추출)
  const locationList = useMemo(() => {
    if (tableData.length === 0 || !locationKey) return [];
    
    // 선택된 장치가 '전체'가 아닐 때는 해당 장치의 행만 대상으로 개소 목록 추출
    const targetRows = (selectedDevice !== "전체" && deviceKey)
      ? tableData.filter(row => String(row[deviceKey] || '').trim() === selectedDevice)
      : tableData;

    const locations = new Set<string>();
    targetRows.forEach(row => {
      const val = row[locationKey];
      if (val) {
        const cleanVal = String(val).trim();
        if (cleanVal !== '') locations.add(cleanVal);
      }
    });
    return Array.from(locations).sort((a, b) => a.localeCompare(b, 'ko'));
  }, [tableData, locationKey, selectedDevice, deviceKey]);

  // 2. 장치별 발생 수량 차트 데이터
  const chartData = useMemo(() => {
    if (tableData.length === 0 || !deviceKey) return [];
    const counts: Record<string, number> = {};
    tableData.forEach(row => {
      const deviceVal = String(row[deviceKey] || '기타').trim();
      if (deviceVal && deviceVal.length < 20) {
        counts[deviceVal] = (counts[deviceVal] || 0) + 1;
      }
    });
    return Object.entries(counts)
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count);
  }, [tableData, deviceKey]);

  const maxCount = useMemo(() => {
    if (chartData.length === 0) return 1;
    return Math.max(...chartData.map(item => item.count));
  }, [chartData]);

  // 3. 필터링 및 날짜순 정렬된 테이블 데이터
  const filteredTableData = useMemo(() => {
    let result = [...tableData];

    // 장치 선택 필터
    if (selectedDevice !== "전체" && deviceKey) {
      result = result.filter(row => String(row[deviceKey] || '').trim() === selectedDevice);
    }

    // 개소 선택 필터
    if (selectedLocation !== "전체" && locationKey) {
      result = result.filter(row => String(row[locationKey] || '').trim() === selectedLocation);
    }

    // 세부작업내용 키워드 검색 필터
    if (searchText.trim() !== "" && workContentKey) {
      const keyword = searchText.toLowerCase();
      result = result.filter(row => String(row[workContentKey] || '').toLowerCase().includes(keyword));
    }

    // 발생일(날짜) 기준 최신순(내림차순) 자동 정렬
    if (dateKey) {
      result.sort((a, b) => {
        const dateA = String(a[dateKey] || '');
        const dateB = String(b[dateKey] || '');
        return dateB.localeCompare(dateA);
      });
    }

    return result;
  }, [tableData, selectedDevice, selectedLocation, searchText, deviceKey, locationKey, workContentKey, dateKey]);

  // 4. 스마트 요약 및 반복 작업 감지
  const handleSummary = () => {
    if (showSummary) {
      setShowSummary(false);
      return;
    }

    setIsSummarizing(true);
    setTimeout(() => {
      const targetData = filteredTableData;
      const locationCounts: Record<string, number> = {};
      const workCounts: Record<string, number> = {};

      targetData.forEach(row => {
        const loc = locationKey ? String(row[locationKey] || '기타 개소').trim() : '전체 개소';
        locationCounts[loc] = (locationCounts[loc] || 0) + 1;

        if (workContentKey) {
          const content = String(row[workContentKey] || '').trim();
          if (content) {
            // 앞뒤 공백 및 유사 텍스트 집계
            const normalized = content.length > 40 ? content.substring(0, 40) + '...' : content;
            workCounts[normalized] = (workCounts[normalized] || 0) + 1;
          }
        }
      });

      const locationSummaryText = Object.entries(locationCounts)
        .sort((a, b) => b[1] - a[1])
        .map(([loc, count]) => `   - [개소: ${loc}] : ${count}건`)
        .join('\n');

      const repeatedWorks = Object.entries(workCounts)
        .filter(([_, count]) => count > 1)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 6);

      const repeatedWorksText = repeatedWorks.length > 0 
        ? repeatedWorks.map(([work, count]) => `   ⚠️ "${work}" ➔ ${count}회 반복 발생`).join('\n')
        : '   - 2회 이상 반복된 동일/유사 작업 내용이 없습니다.';

      setSummary(`📊 [유지보수 실적 스마트 요약 및 반복작업 감지 리포트]
- 분석 파일: ${fileName}
- 필터 조건: 장치 [ ${selectedDevice} ], 개소 [ ${selectedLocation} ] (조회된 건수: 총 ${targetData.length}건)

📌 [지역/개소별 실적 분포]
${locationSummaryText || '   - 데이터 없음'}

🔄 [동일 및 유사 보수실적 반복 감지 분석]
${repeatedWorksText}`);
      
      setIsSummarizing(false);
      setShowSummary(true);
    }, 300);
  };

  // 5. 핵심 4개 컬럼([장치명], [개소], [세부작업내용], [발생일]) 엑셀 내보내기
  const handleExportExcel = () => {
    if (filteredTableData.length === 0) {
      alert('내보낼 데이터가 없습니다.');
      return;
    }

    // 4가지 핵심 컬럼만 추출하여 정제된 객체 배열 생성
    const exportRows = filteredTableData.map(row => ({
      '장치명': (deviceKey && row[deviceKey]) ? String(row[deviceKey]).trim() : '',
      '개소': (locationKey && row[locationKey]) ? String(row[locationKey]).trim() : '',
      '세부작업내용': (workContentKey && row[workContentKey]) ? String(row[workContentKey]).trim() : '',
      '발생일': (dateKey && row[dateKey]) ? String(row[dateKey]).trim() : ''
    }));

    // 워크북 생성 및 파일 저장
    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    
    // 열 너비 자동 설정 (가독성 향상)
    worksheet['!cols'] = [
      { wch: 18 }, // 장치명
      { wch: 20 }, // 개소
      { wch: 55 }, // 세부작업내용
      { wch: 14 }  // 발생일
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '유지보수실적_핵심4개컬럼');

    const cleanDevicePrefix = selectedDevice !== '전체' ? `${selectedDevice}_` : '';
    const dateStr = new Date().toISOString().slice(0, 10);
    const exportFileName = `유지보수실적_${cleanDevicePrefix}${dateStr}.xlsx`;

    XLSX.writeFile(workbook, exportFileName);
  };

  // 6. 핵심 4개 컬럼 및 리포트 PDF 내보내기 (브라우저 인쇄/PDF 저장 팝업)
  const handleExportPdf = () => {
    if (filteredTableData.length === 0) {
      alert('PDF로 내보낼 데이터가 없습니다.');
      return;
    }

    const printWindow = window.open('', '_blank');
    if (!printWindow) {
      alert('팝업 차단이 설정되어 있습니다. 팝업 허용 후 다시 시도해 주세요.');
      return;
    }

    const todayStr = new Date().toLocaleDateString('ko-KR', {
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });

    const rowsHtml = filteredTableData.map((row, idx) => `
      <tr>
        <td style="text-align: center; color: #64748b; font-size: 8.5pt;">${idx + 1}</td>
        <td style="font-weight: 700; color: #1e293b; white-space: nowrap;">${(deviceKey && row[deviceKey]) ? String(row[deviceKey]).trim() : '-'}</td>
        <td style="color: #334155; white-space: nowrap; text-align: center;">${(locationKey && row[locationKey]) ? String(row[locationKey]).trim() : '-'}</td>
        <td style="color: #0f172a; word-break: break-all; white-space: normal; line-height: 1.4;">${(workContentKey && row[workContentKey]) ? String(row[workContentKey]).trim() : '-'}</td>
        <td style="text-align: center; color: #475569; font-family: monospace; white-space: nowrap;">${(dateKey && row[dateKey]) ? String(row[dateKey]).trim() : '-'}</td>
      </tr>
    `).join('');

    const htmlContent = `
      <!DOCTYPE html>
      <html lang="ko">
      <head>
        <meta charset="UTF-8" />
        <title>유지보수실적_보고서_${selectedDevice !== '전체' ? selectedDevice + '_' : ''}${new Date().toISOString().slice(0, 10)}</title>
        <style>
          @page {
            size: A4 portrait;
            margin: 12mm 10mm;
          }
          * {
            box-sizing: border-box;
            font-family: -apple-system, BlinkMacSystemFont, "Malgun Gothic", "맑은 고딕", "Apple SD Gothic Neo", sans-serif;
          }
          body {
            margin: 0;
            padding: 10px;
            color: #0f172a;
            background: #ffffff;
            font-size: 10pt;
            line-height: 1.5;
          }
          .header {
            border-bottom: 2px solid #2563eb;
            padding-bottom: 10px;
            margin-bottom: 16px;
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
          }
          .title {
            font-size: 16pt;
            font-weight: 800;
            color: #1e293b;
            margin: 0;
          }
          .sub-title {
            font-size: 9.5pt;
            color: #2563eb;
            font-weight: 600;
            margin-top: 3px;
          }
          .meta-info {
            font-size: 8.5pt;
            color: #64748b;
            text-align: right;
            line-height: 1.4;
          }
          .summary-card {
            background-color: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 10px 14px;
            margin-bottom: 14px;
            font-size: 9pt;
          }
          .summary-card table {
            width: 100%;
            border-collapse: collapse;
          }
          .summary-card td {
            padding: 3px 6px;
          }
          .summary-title {
            font-weight: 700;
            color: #1e40af;
            margin-bottom: 4px;
            font-size: 9.5pt;
          }
          table.data-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 8.5pt;
            margin-top: 8px;
          }
          table.data-table th, table.data-table td {
            border: 1px solid #cbd5e1;
            padding: 6px 7px;
          }
          table.data-table th {
            background-color: #f1f5f9;
            color: #334155;
            font-weight: 700;
            text-align: center;
          }
          table.data-table tr:nth-child(even) {
            background-color: #f8fafc;
          }
          .footer {
            margin-top: 20px;
            text-align: center;
            font-size: 8pt;
            color: #94a3b8;
            border-top: 1px solid #e2e8f0;
            padding-top: 6px;
          }
          @media print {
            body { padding: 0; }
            .no-print { display: none; }
            table.data-table { page-break-inside: auto; }
            tr { page-break-inside: avoid; page-break-after: auto; }
            thead { display: table-header-group; }
          }
        </style>
      </head>
      <body>
        <div class="header">
          <div>
            <h1 class="title">📋 실시간 유지보수 실적 보고서</h1>
            <div class="sub-title">
              파일: ${fileName} (추출 건수: 총 ${filteredTableData.length}건)
            </div>
          </div>
          <div class="meta-info">
            <div><strong>출력 일자:</strong> ${todayStr}</div>
            <div><strong>필터 조건:</strong> 장치 [ ${selectedDevice} ] / 개소 [ ${selectedLocation} ]</div>
          </div>
        </div>

        <div class="summary-card">
          <div class="summary-title">📌 필터 및 조회 요약</div>
          <table>
            <tr>
              <td width="25%"><strong>선택 장치:</strong> ${selectedDevice}</td>
              <td width="25%"><strong>선택 개소:</strong> ${selectedLocation}</td>
              <td width="25%"><strong>검색 키워드:</strong> ${searchText ? `"${searchText}"` : '전체'}</td>
              <td width="25%"><strong>추출 건수:</strong> 총 ${filteredTableData.length}건</td>
            </tr>
          </table>
        </div>

        <table class="data-table">
          <thead>
            <tr>
              <th style="width: 32px;">No</th>
              <th style="width: 100px;">장치명</th>
              <th style="width: 100px;">개소</th>
              <th>세부작업내용</th>
              <th style="width: 80px;">발생일</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>

        <div class="footer">
          철도/KTX 실시간 유지보수 실적 데이터 분석 웹앱에서 생성된 공식 리포트입니다.
        </div>

        <script>
          window.onload = function() {
            setTimeout(function() {
              window.print();
            }, 300);
          };
        </script>
      </body>
      </html>
    `;

    printWindow.document.open();
    printWindow.document.write(htmlContent);
    printWindow.document.close();
  };

  // 요약 리포트 클립보드 복사
  const handleCopySummary = () => {
    if (!summary) return;
    navigator.clipboard.writeText(summary);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <main className="min-h-screen bg-slate-900 text-slate-100 py-8 px-4 sm:px-6 lg:px-8 font-sans">
      <div className="max-w-7xl mx-auto space-y-6">
        
        {/* 상단 대시보드 타이틀 헤더 */}
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center bg-slate-800/90 backdrop-blur p-6 rounded-2xl shadow-xl border border-slate-700 gap-4">
          <div className="flex items-center space-x-3.5">
            <div className="p-3 bg-gradient-to-br from-blue-600 to-indigo-600 rounded-xl text-white shadow-lg">
              <Database className="w-7 h-7" />
            </div>
            <div>
              <h1 className="text-xl md:text-2xl font-black tracking-tight text-white flex items-center gap-2">
                실시간 유지보수 실적 데이터 분석 웹앱
              </h1>
              <p className="text-xs text-slate-400 mt-0.5">
                분석 파일: <span className="text-cyan-400 font-semibold">{fileName}</span> | 
                인식된 장치: <strong className="text-emerald-400 font-bold">{deviceList.length}종류</strong> | 
                총 데이터: <strong className="text-white">{tableData.length}건</strong>
              </p>
            </div>
          </div>
          
          <div className="flex items-center space-x-2.5 flex-wrap gap-y-2">
            {/* 요약하기 토글 버튼 */}
            <button
              onClick={handleSummary}
              disabled={isSummarizing}
              className={`flex items-center space-x-2 px-4 py-2 rounded-xl font-bold shadow-md transition-all active:scale-95 text-xs md:text-sm ${
                showSummary 
                  ? 'bg-slate-700 hover:bg-slate-600 text-slate-200 border border-slate-600' 
                  : 'bg-indigo-600 hover:bg-indigo-500 text-white shadow-indigo-600/30'
              }`}
            >
              {showSummary ? <EyeOff className="w-4 h-4 text-slate-300" /> : <FileText className="w-4 h-4 text-indigo-200" />}
              <span>{isSummarizing ? '분석 중...' : showSummary ? '📊 요약 숨기기' : '📊 요약하기'}</span>
            </button>

            {/* 파일 업로드 버튼 */}
            <label className="cursor-pointer bg-blue-600 hover:bg-blue-500 text-white text-xs md:text-sm px-4 py-2 rounded-xl font-bold shadow-md shadow-blue-600/30 transition-all flex items-center space-x-1.5 active:scale-95">
              <UploadCloud className="w-4 h-4" />
              <span>파일 업로드 (.xlsx / .csv)</span>
              <input
                type="file"
                accept=".xlsx, .xls, .csv, .txt"
                onChange={handleFileUpload}
                className="hidden"
              />
            </label>
          </div>
        </div>

        {/* 4. 스마트 요약 및 반복작업 감지 리포트 박스 (토글) */}
        {showSummary && summary && (
          <div className="bg-indigo-950/70 border border-indigo-500/40 p-6 rounded-2xl shadow-2xl relative transition-all duration-300 animate-in fade-in">
            <div className="flex justify-between items-center mb-3">
              <div className="flex items-center space-x-2 text-indigo-300 font-bold text-sm md:text-base">
                <FileText className="w-5 h-5 text-amber-400" />
                <span>스마트 요약 및 동일/유사 보수실적 반복 감지 리포트</span>
              </div>
              
              <div className="flex items-center space-x-2">
                <button
                  onClick={handleCopySummary}
                  className="text-xs bg-indigo-900/80 hover:bg-indigo-800 text-indigo-200 px-3 py-1.5 rounded-lg border border-indigo-400/30 transition-colors flex items-center space-x-1 font-medium"
                >
                  {copied ? <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                  <span>{copied ? '복사됨!' : '리포트 복사'}</span>
                </button>
                <button 
                  onClick={() => setShowSummary(false)}
                  className="text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 px-3 py-1.5 rounded-lg border border-slate-600 transition-colors flex items-center space-x-1"
                >
                  <EyeOff className="w-3.5 h-3.5" />
                  <span>숨기기</span>
                </button>
              </div>
            </div>

            <div className="bg-slate-950/60 p-4 rounded-xl border border-indigo-900/50">
              <p className="text-indigo-100 whitespace-pre-line leading-relaxed font-mono text-xs md:text-sm">
                {summary}
              </p>
            </div>
          </div>
        )}

        {/* 2. 장치명 시각화 막대 차트 및 선택 필터 (B열 기준) */}
        <div className="bg-slate-800 p-6 rounded-2xl shadow-xl border border-slate-700 space-y-5">
          <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 border-b border-slate-700/80 pb-4">
            <div>
              <h2 className="text-base md:text-lg font-bold flex items-center space-x-2 text-slate-100">
                <BarChart2 className="w-5 h-5 text-emerald-400" />
                <span>장치명별 발생 수량 시각화 차트 (B열 기준)</span>
              </h2>
              <p className="text-xs text-slate-400 mt-1 flex items-center">
                <MousePointerClick className="w-3.5 h-3.5 mr-1 text-cyan-400" /> 
                차트의 <b>막대나 건수를 클릭</b>하면 하단 분석 테이블이 해당 장치로 실시간 필터링됩니다.
              </p>
            </div>
            
            {/* 장치명 가나다순 드롭다운 필터 */}
            <div className="flex items-center space-x-2 bg-slate-900 px-3.5 py-1.5 rounded-xl border border-slate-700 shadow-inner">
              <ListFilter className="w-4 h-4 text-cyan-400" />
              <span className="text-xs text-slate-300 font-semibold">장치 필터:</span>
              <select 
                value={selectedDevice} 
                onChange={(e) => handleDeviceSelect(e.target.value)}
                className="bg-slate-800 text-cyan-300 font-bold text-xs rounded-lg px-2.5 py-1 outline-none border border-slate-600 cursor-pointer max-w-[200px]"
              >
                <option value="전체">✨ 전체 장치 보기 ({deviceList.length}개)</option>
                {deviceList.map(device => (
                  <option key={device} value={device}>{device}</option>
                ))}
              </select>
            </div>
          </div>

          {/* 수량 막대 차트 영역 */}
          <div className="bg-slate-900/80 p-5 rounded-xl border border-slate-700/60 space-y-2.5 max-h-80 overflow-y-auto">
            {chartData.length > 0 ? (
              chartData.map((item) => {
                const percentage = Math.round((item.count / maxCount) * 100);
                const isSelected = selectedDevice === item.name;
                
                return (
                  <div 
                    key={item.name} 
                    onClick={() => handleDeviceSelect(isSelected ? "전체" : item.name)}
                    className={`group space-y-1.5 p-2.5 rounded-xl transition-all cursor-pointer ${
                      isSelected 
                        ? 'bg-blue-950/70 border-2 border-blue-500 shadow-lg shadow-blue-950/50' 
                        : 'hover:bg-slate-800/70 border border-transparent hover:border-slate-700'
                    }`}
                  >
                    <div className="flex justify-between items-center text-xs md:text-sm flex-wrap gap-2">
                      <span className={`font-bold transition-colors ${isSelected ? 'text-blue-300' : 'text-cyan-300 group-hover:text-cyan-200'}`}>
                        {item.name} {isSelected && <span className="text-xs bg-blue-600 text-white px-2 py-0.5 rounded-full ml-1.5">선택됨</span>}
                      </span>
                      
                      <div className="flex items-center space-x-2.5">
                        {/* [UI/UX 개선] 장치 선택 시 해당 장치 막대 내부에 개소 선택 드롭다운 직접 배치 */}
                        {isSelected && locationList.length > 0 && (
                          <div 
                            onClick={(e) => e.stopPropagation()} 
                            className="flex items-center space-x-1.5 bg-slate-900/90 border border-cyan-400/80 px-2.5 py-1 rounded-lg shadow-md animate-in fade-in"
                          >
                            <MapPin className="w-3.5 h-3.5 text-cyan-300 animate-bounce" />
                            <span className="text-xs text-cyan-200 font-bold whitespace-nowrap">개소 선택:</span>
                            <select 
                              value={selectedLocation} 
                              onChange={(e) => setSelectedLocation(e.target.value)}
                              className="bg-slate-800 text-cyan-200 font-bold text-xs rounded-md px-2 py-0.5 outline-none border border-cyan-400/50 cursor-pointer hover:border-cyan-300 transition-colors"
                            >
                              <option value="전체">✨ 전체 개소 보기 ({locationList.length}개소)</option>
                              {locationList.map(loc => (
                                <option key={loc} value={loc}>{loc}</option>
                              ))}
                            </select>
                          </div>
                        )}

                        <span className={`px-2.5 py-0.5 rounded-full text-xs font-mono font-bold transition-colors ${
                          isSelected 
                            ? 'bg-blue-600 text-white shadow' 
                            : 'bg-slate-800 text-cyan-300 border border-slate-700 group-hover:bg-blue-600 group-hover:text-white'
                        }`}>
                          {item.count} 건
                        </span>
                      </div>
                    </div>
                    
                    <div className="w-full bg-slate-800 h-3.5 rounded-full overflow-hidden p-0.5 border border-slate-700/80">
                      <div 
                        className={`h-full rounded-full transition-all duration-500 ${
                          isSelected ? 'bg-gradient-to-r from-cyan-400 to-blue-500' : 'bg-gradient-to-r from-blue-600 to-indigo-500 group-hover:from-cyan-500 group-hover:to-blue-500'
                        }`}
                        style={{ width: `${Math.max(percentage, 6)}%` }}
                      ></div>
                    </div>
                  </div>
                );
              })
            ) : (
              <p className="text-center py-6 text-slate-500 text-sm">표시할 장치 데이터가 없습니다.</p>
            )}
          </div>
          
          {selectedDevice !== "전체" && (
            <div className="flex justify-between items-center bg-blue-950/40 border border-blue-500/30 p-3 rounded-xl">
              <span className="text-xs text-blue-200">
                현재 <b>[ {selectedDevice} ]</b> 장치 필터가 적용되어 있습니다. (총 {filteredTableData.length}건)
              </span>
              <button 
                onClick={() => handleDeviceSelect("전체")}
                className="text-xs bg-slate-800 hover:bg-slate-700 text-cyan-300 px-3 py-1.5 rounded-lg font-semibold border border-slate-600 transition-colors flex items-center gap-1"
              >
                <RefreshCw className="w-3.5 h-3.5" />
                <span>필터 초기화</span>
              </button>
            </div>
          )}
        </div>

        {/* 3. 날짜별 분석 데이터 테이블 및 5. 핵심 4개 컬럼 엑셀 내보내기 구역 */}
        <div className="bg-slate-800 p-6 rounded-2xl shadow-xl border border-slate-700 space-y-4">
          <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4">
            
            {/* 타이틀 및 정렬 배지 */}
            <div className="flex items-center space-x-3 flex-wrap gap-2">
              <h2 className="text-base md:text-lg font-extrabold flex items-center space-x-2 text-white">
                <Database className="w-5 h-5 text-blue-400" />
                <span>분석 데이터 내용 확인 ({filteredTableData.length}행)</span>
              </h2>

              <span className="text-xs bg-emerald-500/20 text-emerald-300 border border-emerald-500/40 px-2.5 py-1 rounded-full flex items-center font-medium">
                <Calendar className="w-3 h-3 mr-1" /> 발생일 최신순(내림차순) 정렬
              </span>
            </div>
            
            {/* 우측 컨트롤 바: 세부작업내용 검색, 엑셀/PDF 내보내기 버튼 */}
            <div className="flex items-center space-x-2.5 flex-wrap gap-2 w-full lg:w-auto">
              
              {/* 세부작업내용 실시간 키워드 검색 */}
              <div className="flex items-center space-x-2 bg-slate-900 px-3 py-1.5 rounded-xl border border-slate-700 flex-1 sm:flex-initial">
                <Search className="w-4 h-4 text-cyan-400" />
                <input 
                  type="text"
                  placeholder="세부작업내용 실시간 검색..."
                  value={searchText}
                  onChange={(e) => setSearchText(e.target.value)}
                  className="bg-transparent text-cyan-200 text-xs outline-none w-full sm:w-44 placeholder:text-slate-500"
                />
              </div>

              {/* 5. 핵심 4개 컬럼 엑셀 내보내기 버튼 */}
              <button
                onClick={handleExportExcel}
                className="bg-emerald-600 hover:bg-emerald-500 text-white text-xs px-3.5 py-2 rounded-xl font-bold shadow-md shadow-emerald-600/30 transition-all flex items-center space-x-1.5 active:scale-95 whitespace-nowrap"
                title="[장치명], [개소], [세부작업내용], [발생일] 4개 컬럼만 추출하여 엑셀 다운로드"
              >
                <Download className="w-4 h-4" />
                <span>📥 엑셀 내보내기 (4개 컬럼)</span>
              </button>

              {/* 6. 핵심 4개 컬럼 PDF 내보내기 버튼 */}
              <button
                onClick={handleExportPdf}
                className="bg-rose-600 hover:bg-rose-500 text-white text-xs px-3.5 py-2 rounded-xl font-bold shadow-md shadow-rose-600/30 transition-all flex items-center space-x-1.5 active:scale-95 whitespace-nowrap"
                title="현재 필터링된 데이터를 깔끔한 A4 PDF 보고서로 인쇄/저장"
              >
                <Printer className="w-4 h-4" />
                <span>📄 PDF로 내보내기</span>
              </button>
            </div>
          </div>

          {/* 데이터 테이블 */}
          <div className="w-full overflow-x-auto border border-slate-700 rounded-xl bg-slate-900 shadow-inner">
            <div className="max-h-[500px] overflow-y-auto">
              <table className="w-full border-collapse text-left text-xs md:text-sm">
                <thead className="bg-slate-800 text-slate-300 uppercase font-semibold sticky top-0 z-10 shadow border-b border-slate-700">
                  <tr>
                    <th className="px-4 py-3 text-center w-12 text-slate-400">No.</th>
                    {columns.map((col) => {
                      const isLocation = col === locationKey;
                      const isDate = col === dateKey;
                      return (
                        <th 
                          key={col} 
                          className={`px-4 py-3 whitespace-nowrap ${isLocation || isDate ? 'text-center' : ''}`}
                        >
                          {col}
                        </th>
                      );
                    })}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800 text-slate-300">
                  {filteredTableData.length > 0 ? (
                    filteredTableData.map((row, index) => (
                      <tr key={index} className="hover:bg-slate-800/60 transition-colors">
                        <td className="px-4 py-3 text-center text-slate-500 font-mono text-xs">
                          {index + 1}
                        </td>
                        {columns.map((col) => {
                          const isWorkContent = col.includes('세부작업') || col.includes('작업내용') || col.includes('내용') || col.includes('현상');
                          const isDevice = col === deviceKey;
                          const isDate = col === dateKey;
                          const isLocation = col === locationKey;

                          return (
                            <td 
                              key={col} 
                              className={`px-4 py-3 ${
                                isWorkContent 
                                  ? 'min-w-[280px] max-w-xl text-cyan-100 font-normal whitespace-normal break-words leading-relaxed' 
                                  : isDevice
                                  ? 'whitespace-nowrap font-bold text-cyan-300'
                                  : isDate
                                  ? 'whitespace-nowrap font-mono text-emerald-300 text-center'
                                  : isLocation
                                  ? 'whitespace-nowrap text-rose-300 text-center font-medium'
                                  : 'whitespace-nowrap max-w-xs truncate'
                              }`}
                            >
                              {String(row[col] || '')}
                            </td>
                          );
                        })}
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={columns.length + 1} className="text-center py-16 text-slate-500">
                        선택하신 조건(장치/개소/세부작업내용 검색)에 부합하는 데이터가 없습니다.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
          
          <div className="flex flex-col sm:flex-row justify-between items-center text-xs text-slate-400 gap-2 pt-1">
            <p>* <b>[📥 엑셀 내보내기 (4개 컬럼)]</b> 버튼을 클릭하면 불필요한 열을 제외하고 <b>[장치명, 개소, 세부작업내용, 발생일]</b> 4개 컬럼만 정제되어 즉시 다운로드됩니다.</p>
            <p className="text-slate-500 font-mono">표시 중인 데이터: {filteredTableData.length}건</p>
          </div>
        </div>

      </div>
    </main>
  );
}