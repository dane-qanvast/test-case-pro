-- ─────────────────────────────────────────────────────────────────────────────
-- Per-tester findings (Nov 2026): every tester keeps their own result for every
-- test, instead of all testers sharing (and overwriting) one result per test.
--
-- What this does:
--   1. Adds a tester_name column to findings.
--   2. Files each existing finding under whoever last edited it, so nothing is lost.
--   3. Changes the key from "one row per test" to "one row per test per tester".
--
-- Safe to run more than once: each step checks whether it has already been done.
-- ─────────────────────────────────────────────────────────────────────────────
alter table findings add column if not exists tester_name text;

update findings
   set tester_name = coalesce(nullif(updated_by, ''), nullif(tested_by, ''), '')
 where tester_name is null;

alter table findings alter column tester_name set default '';
alter table findings alter column tester_name set not null;

do $$
declare
  pk_name text;
  pk_cols int;
begin
  select c.conname, array_length(c.conkey, 1)
    into pk_name, pk_cols
    from pg_constraint c
   where c.conrelid = 'public.findings'::regclass and c.contype = 'p';
  if pk_cols is distinct from 2 then
    if pk_name is not null then
      execute format('alter table findings drop constraint %I', pk_name);
    end if;
    alter table findings add constraint findings_pkey primary key (test_case_id, tester_name);
  end if;
end $$;

create index if not exists idx_findings_tester on findings(session_id, tester_name);
