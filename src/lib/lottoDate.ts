/**
 * 동행복권(로또 6/45) 회차 및 날짜 계산 유틸리티
 */

// 로또 1회차 추첨일: 2002년 12월 7일 (토)
// 추첨 방송 시작 및 번호 공개 가능 시각: 매주 토요일 오후 8시 45분 (20:45) KST
// 한국 표준시(KST)는 UTC+9
// UTC 기준 시점: 2002-12-07 11:45:00 UTC (11시 45분)
const LOTTO_BASE_UTC_MS = Date.UTC(2002, 11, 7, 11, 45, 0);
const ONE_WEEK_MS = 7 * 24 * 60 * 60 * 1000;

/**
 * 한국 표준시(KST) 기준 현재 날짜/시각을 기반으로 예상 최신 로또 회차를 계산합니다.
 * - 매주 토요일 20시 45분 이전: 직전 주 토요일 회차
 * - 매주 토요일 20시 45분 이후: 당일 추첨 회차
 * @param now 기준 시각 (기본값: 현재 시각)
 * @returns 예상 최신 회차 번호
 */
export function getEstimatedLatestDrawNo(now: Date = new Date()): number {
  const currentUtcMs = now.getTime();
  const diffMs = currentUtcMs - LOTTO_BASE_UTC_MS;
  if (diffMs < 0) {
    return 1;
  }
  return 1 + Math.floor(diffMs / ONE_WEEK_MS);
}

/**
 * 8자리 날짜 문자열(YYYYMMDD)을 YYYY.MM.DD 형식으로 포맷팅합니다.
 */
export function formatLottoDate(ymd?: string): string {
  if (!ymd || ymd.length < 8) return '';
  const clean = ymd.replace(/[^0-9]/g, '');
  if (clean.length === 8) {
    return `${clean.slice(0, 4)}.${clean.slice(4, 6)}.${clean.slice(6, 8)}`;
  }
  return ymd;
}
