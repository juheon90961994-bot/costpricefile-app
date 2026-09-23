-- cost 테이블의 단가 컬럼에서 8.5, 18.5, 12.5 같은 소수값을
-- 반올림 없이 저장할 수 있도록 integer 계열을 numeric으로 변경합니다.
-- Supabase Dashboard > SQL Editor에서 한 번만 실행하세요.

begin;

alter table public.cost
  alter column "달러단가" type numeric
  using nullif(trim("달러단가"::text), '')::numeric,
  alter column "원화단가" type numeric
  using nullif(trim("원화단가"::text), '')::numeric;

commit;

-- 변경 결과 확인용입니다. 두 행의 data_type이 numeric이면 정상입니다.
select column_name, data_type, numeric_precision, numeric_scale
from information_schema.columns
where table_schema = 'public'
  and table_name = 'cost'
  and column_name in ('달러단가', '원화단가')
order by column_name;
