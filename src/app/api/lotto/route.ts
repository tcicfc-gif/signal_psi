import { NextResponse } from 'next/server';
import { getEstimatedLatestDrawNo, formatLottoDate } from '@/lib/lottoDate';

// Vercel ISR 및 동적 렌더링 설정
// searchParams 사용을 위해 force-dynamic 적용 + fetch 캐싱(3600초) 및 revalidate(1800초)
export const dynamic = 'force-dynamic';
export const revalidate = 1800;

export interface LottoItem {
  drwNo: number;
  drwNoDate: string;
  numbers: number[];
  bonusNo: number;
  firstWinamnt: number;
  firstPrzwnerCo: number;
  firstAccumamnt: number;
  totSellamnt: number;
}

const COMMON_HEADERS = {
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
  'Referer': 'https://www.dhlottery.co.kr/',
  'Accept': 'application/json, text/javascript, */*; q=0.01'
};

/**
 * 신규 동행복권 API 응답 로우 데이터를 LottoItem 규격으로 변환
 */
function mapRawToLottoItem(raw: any): LottoItem {
  const drwNo = Number(raw.ltEpsd ?? raw.drwNo ?? 0);
  const rawDate = String(raw.ltRflYmd ?? raw.drwNoDate ?? '');
  const drwNoDate = formatLottoDate(rawDate);

  const numbers = [
    Number(raw.tm1WnNo ?? raw.drwtNo1 ?? 0),
    Number(raw.tm2WnNo ?? raw.drwtNo2 ?? 0),
    Number(raw.tm3WnNo ?? raw.drwtNo3 ?? 0),
    Number(raw.tm4WnNo ?? raw.drwtNo4 ?? 0),
    Number(raw.tm5WnNo ?? raw.drwtNo5 ?? 0),
    Number(raw.tm6WnNo ?? raw.drwtNo6 ?? 0),
  ].sort((a, b) => a - b);

  const bonusNo = Number(raw.bnsWnNo ?? raw.bnusNo ?? 0);
  const firstWinamnt = Number(raw.rnk1WnAmt ?? raw.firstWinamnt ?? 0);
  const firstPrzwnerCo = Number(raw.rnk1WnNope ?? raw.firstPrzwnerCo ?? 0);
  const firstAccumamnt = Number(raw.rnk1SumWnAmt ?? raw.firstAccumamnt ?? 0);
  const totSellamnt = Number(raw.rlvtEpsdSumNtslAmt ?? raw.totSellamnt ?? 0);

  return {
    drwNo,
    drwNoDate,
    numbers,
    bonusNo,
    firstWinamnt,
    firstPrzwnerCo,
    firstAccumamnt,
    totSellamnt,
  };
}

/**
 * 특정 단일 회차 조회 (Fail-safe 검증용)
 */
async function fetchSingleDraw(drawNo: number): Promise<LottoItem | null> {
  // 1차: 신규 API 엔드포인트
  const newUrl = `https://www.dhlottery.co.kr/lt645/selectPstLt645Info.do?srchLtEpsd=${drawNo}`;
  try {
    const res = await fetch(newUrl, {
      headers: COMMON_HEADERS,
      next: { revalidate: 3600 }
    });
    if (res.ok) {
      const json = await res.json();
      const list = json?.data?.list;
      if (Array.isArray(list) && list.length > 0 && list[0].ltEpsd) {
        return mapRawToLottoItem(list[0]);
      }
    }
  } catch (err) {
    console.warn(`신규 API 단일 회차(${drawNo}) 조회 에러:`, err);
  }

  // 2차: 레거시 API 엔드포인트 fallback
  const legacyUrl = `https://www.dhlottery.co.kr/common.do?method=getLottoNumber&drwNo=${drawNo}`;
  try {
    const res = await fetch(legacyUrl, {
      headers: COMMON_HEADERS,
      next: { revalidate: 3600 }
    });
    if (res.ok) {
      const json = await res.json();
      if (json && json.returnValue === 'success') {
        return mapRawToLottoItem(json);
      }
    }
  } catch (err) {
    console.warn(`레거시 API 단일 회차(${drawNo}) 조회 에러:`, err);
  }

  return null;
}

/**
 * 전체 회차 일괄 조회 (동행복권 단일 요청 최적화)
 */
async function fetchAllDraws(): Promise<LottoItem[] | null> {
  const url = 'https://www.dhlottery.co.kr/lt645/selectPstLt645Info.do?srchLtEpsd=all';
  try {
    const res = await fetch(url, {
      headers: COMMON_HEADERS,
      next: { revalidate: 3600 }
    });
    if (res.ok) {
      const json = await res.json();
      const list = json?.data?.list;
      if (Array.isArray(list) && list.length > 0) {
        return list.map(mapRawToLottoItem);
      }
    }
  } catch (err) {
    console.warn('전체 회차(all) 일괄 조회 실패, 개별 수집으로 전환:', err);
  }
  return null;
}

/**
 * GET /api/lotto
 * 
 * Query Parameters:
 * - count: 가져올 최근 회차 수 (기본: 100, 최소: 1, 최대: 100)
 */
export async function GET(request: Request) {
  try {
    const { searchParams } = new URL(request.url);
    const countParam = parseInt(searchParams.get('count') || '100', 10);
    const targetCount = isNaN(countParam) ? 100 : Math.min(Math.max(1, countParam), 100);

    // 1. 현재 시각(KST) 기준 예상 최신 회차 번호 계산
    const estimatedDrawNo = getEstimatedLatestDrawNo();

    // 2. 최신 회차 API 호출 및 안전 검증 (Fail-safe)
    let latestItem = await fetchSingleDraw(estimatedDrawNo);
    let confirmedLatestNo = estimatedDrawNo;

    // 만약 토요일 밤 아직 추첨 결과가 서버에 등록되지 않은 경우 직전 회차로 자동 확정
    if (!latestItem) {
      confirmedLatestNo = Math.max(1, estimatedDrawNo - 1);
      latestItem = await fetchSingleDraw(confirmedLatestNo);
    }

    // 3. 최신 N개 데이터 수집
    let results: LottoItem[] = [];

    // 방법 A: 일괄(all) 엔드포인트로 초고속 1회 통신 시도
    const allDraws = await fetchAllDraws();
    if (allDraws && allDraws.length > 0) {
      // confirmedLatestNo 이하의 데이터만 필터링하여 최신순 정렬
      const validDraws = allDraws
        .filter(d => d.drwNo <= confirmedLatestNo)
        .sort((a, b) => b.drwNo - a.drwNo);

      if (validDraws.length > 0) {
        confirmedLatestNo = validDraws[0].drwNo;
        latestItem = validDraws[0];
        results = validDraws.slice(0, targetCount);
      }
    }

    // 방법 B: 일괄 조회가 실패했을 경우 배치 병렬 호출 fallback
    if (results.length === 0) {
      const drawNosToFetch: number[] = [];
      for (let i = 0; i < targetCount; i++) {
        const no = confirmedLatestNo - i;
        if (no >= 1) drawNosToFetch.push(no);
      }

      if (latestItem) {
        results.push(latestItem);
      }

      const remainingNos = drawNosToFetch.filter(no => no !== confirmedLatestNo);
      const BATCH_SIZE = 15;

      for (let i = 0; i < remainingNos.length; i += BATCH_SIZE) {
        const batch = remainingNos.slice(i, i + BATCH_SIZE);
        const batchResults = await Promise.all(batch.map(no => fetchSingleDraw(no)));
        batchResults.forEach(item => {
          if (item) results.push(item);
        });
      }

      results.sort((a, b) => b.drwNo - a.drwNo);
    }

    if (results.length === 0) {
      return NextResponse.json(
        { error: '로또 데이터를 불러오는 데 실패했습니다.' },
        { status: 502 }
      );
    }

    const latestDrawDate = latestItem?.drwNoDate || results[0]?.drwNoDate || '';

    return NextResponse.json({
      success: true,
      latestDrawNo: confirmedLatestNo,
      latestDrawDate,
      totalCount: results.length,
      data: results,
    });
  } catch (error: any) {
    console.error('Lotto API Route Critical Error:', error);
    return NextResponse.json(
      { error: '서버 내부 오류가 발생했습니다.', message: error?.message },
      { status: 500 }
    );
  }
}
