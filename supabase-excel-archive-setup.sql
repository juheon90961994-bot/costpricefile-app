-- Run once in this project's Supabase SQL Editor.
-- Only the dedicated Excel archive is affected. Existing cost data is unchanged.
BEGIN;
INSERT INTO storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
VALUES ('excel-archive', 'excel-archive', false, 20971520,
 ARRAY['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','application/vnd.ms-excel','application/vnd.ms-excel.sheet.macroEnabled'])
ON CONFLICT (id) DO UPDATE SET file_size_limit=EXCLUDED.file_size_limit,
 allowed_mime_types=EXCLUDED.allowed_mime_types;
DROP POLICY IF EXISTS excel_archive_read ON storage.objects;
CREATE POLICY excel_archive_read ON storage.objects FOR SELECT TO anon, authenticated
 USING (bucket_id='excel-archive');
DROP POLICY IF EXISTS excel_archive_insert ON storage.objects;
CREATE POLICY excel_archive_insert ON storage.objects FOR INSERT TO anon, authenticated
 WITH CHECK (bucket_id='excel-archive');
DROP POLICY IF EXISTS excel_archive_update ON storage.objects;
CREATE POLICY excel_archive_update ON storage.objects FOR UPDATE TO anon, authenticated
 USING (bucket_id='excel-archive') WITH CHECK (bucket_id='excel-archive');
DROP POLICY IF EXISTS excel_archive_delete ON storage.objects;
CREATE POLICY excel_archive_delete ON storage.objects FOR DELETE TO anon, authenticated
 USING (bucket_id='excel-archive');
COMMIT;
