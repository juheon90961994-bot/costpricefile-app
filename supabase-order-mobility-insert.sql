-- 신규 BOM 품목 등록 권한 (기존 자료 변경 없음)
BEGIN;
GRANT INSERT ON TABLE public.order_mobility TO anon, authenticated;
DROP POLICY IF EXISTS order_mobility_public_insert ON public.order_mobility;
CREATE POLICY order_mobility_public_insert ON public.order_mobility FOR INSERT TO anon, authenticated WITH CHECK (true);
COMMIT;
