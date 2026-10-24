-- ============================================================
-- Supabase SQL 마이그레이션: signalpsi 스키마 및 테이블 생성
-- Supabase 대시보드 > SQL Editor에서 이 쿼리를 실행하세요.
-- ============================================================

-- 1) signalpsi 스키마 생성
CREATE SCHEMA IF NOT EXISTS signalpsi;

-- 2) uploaded_files 테이블 (업로드된 파일 메타데이터 관리)
CREATE TABLE IF NOT EXISTS signalpsi.uploaded_files (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  file_name TEXT NOT NULL,
  uploaded_at TIMESTAMPTZ DEFAULT now()
);

-- 3) maintenance_records 테이블 (실제 유지보수 분석 데이터)
CREATE TABLE IF NOT EXISTS signalpsi.maintenance_records (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  file_id UUID NOT NULL REFERENCES signalpsi.uploaded_files(id) ON DELETE CASCADE,
  device_name TEXT DEFAULT '',
  location TEXT DEFAULT '',
  work_content TEXT DEFAULT '',
  occurrence_date TEXT DEFAULT ''
);

-- 4) 인덱스 (조회 성능 향상)
CREATE INDEX IF NOT EXISTS idx_records_file_id ON signalpsi.maintenance_records(file_id);
CREATE INDEX IF NOT EXISTS idx_records_device ON signalpsi.maintenance_records(device_name);

-- 5) RLS(Row Level Security) 활성화 및 공개 접근 정책 설정
ALTER TABLE signalpsi.uploaded_files ENABLE ROW LEVEL SECURITY;
ALTER TABLE signalpsi.maintenance_records ENABLE ROW LEVEL SECURITY;

-- anon/public 키로 전체 CRUD 허용 정책
CREATE POLICY "Allow all for uploaded_files" ON signalpsi.uploaded_files
  FOR ALL USING (true) WITH CHECK (true);

CREATE POLICY "Allow all for maintenance_records" ON signalpsi.maintenance_records
  FOR ALL USING (true) WITH CHECK (true);

-- 6) PostgREST 설정 리로드 알림
-- 주의: Supabase 대시보드 > Settings > API > Exposed schemas에 'signalpsi'를 추가해야 합니다!
NOTIFY pgrst, 'reload config';
