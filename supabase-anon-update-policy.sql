-- 로그인 없이 publishable key로 cost 테이블을 조회하고 수정하기 위한 설정입니다.
-- 주의: publishable key를 가진 모든 사용자가 아래 지정 열을 수정할 수 있습니다.

alter table public.cost enable row level security;

grant select on table public.cost to anon;
grant update (
  "품목",
  "품목명",
  "회사",
  "거래처명",
  "담당자",
  "단위",
  "달러단가",
  "원화단가"
) on table public.cost to anon;
grant insert (
  "품목",
  "품목명",
  "회사",
  "거래처명",
  "담당자",
  "단위",
  "달러단가",
  "원화단가"
) on table public.cost to anon;
grant delete on table public.cost to anon;

drop policy if exists "anon users can read cost" on public.cost;
create policy "anon users can read cost"
on public.cost
for select
to anon
using (true);

drop policy if exists "anon users can update cost" on public.cost;
create policy "anon users can update cost"
on public.cost
for update
to anon
using (true)
with check (true);

drop policy if exists "anon users can insert cost" on public.cost;
create policy "anon users can insert cost"
on public.cost
for insert
to anon
with check (true);

drop policy if exists "anon users can delete cost" on public.cost;
create policy "anon users can delete cost"
on public.cost
for delete
to anon
using (true);
