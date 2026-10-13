-- Test Case Collaborator — Supabase schema
-- Run this once in your Supabase SQL Editor before launching the app.
-- Supabase SQL Editor: left sidebar → SQL Editor → New query → paste this → Run.
-- Safe to re-run — every statement is idempotent.

-- 1. Sessions: one per UAT campaign uploaded by the admin.
create table if not exists sessions (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  created_at timestamptz default now(),
  testers jsonb default '[]'::jsonb,        -- ["Alice", "Bob"]
  developers jsonb default '[]'::jsonb,     -- ["Carol", "Dan"]
  tester_code text not null,
  developer_code text not null,
  tester_col_header text,                   -- the original CSV's tester column header, preserved for export
  base_url text                             -- optional path chip prefix (e.g. https://staging.example.com); testers/devs can override it for themselves
);

-- Idempotent migration for existing deploys that pre-date the base_url
-- column (the create-table above only fires on fresh installs).
alter table sessions add column if not exists base_url text;

-- 2. Test cases: rows from the uploaded CSV.
create table if not exists test_cases (
  id uuid primary key default gen_random_uuid(),
  session_id uuid not null references sessions(id) on delete cascade,
  position integer not null,                -- preserves CSV order
  row_num integer,                          -- original spreadsheet row number for reference
  test_id text,
  module text,
  feature text,
  title text,
  pre_conditions text,
  test_data text,
  steps jsonb default '[]'::jsonb,
  expected_result text,
  downstream_check text,
  edge_case text,
  dev_branch_name text,                     -- optional: git branch where the fix lives, from CSV's "Dev Branch Name" column
  initial_tester text,                      -- pre-assigned tester from the CSV's tester column
  focus text,                               -- optional focus area from the sheet (e.g. Security, Data)
  priority text                             -- optional priority from the sheet (e.g. P0, P1, P2)
);

-- Idempotent migration for existing deploys that pre-date the dev_branch_name
-- column (the create-table above only fires on fresh installs).
alter table test_cases add column if not exists dev_branch_name text;

-- Idempotent migration: Focus area (e.g. Security, Data) and Priority (e.g. P0, P1)
-- imported from the test sheet, used as tester filters.
alter table test_cases add column if not exists focus text;
alter table test_cases add column if not exists priority text;

-- 3. Findings: testers fill these in. One row per test case (upserted as they edit).
create table if not exists findings (
  test_case_id uuid primary key references test_cases(id) on delete cascade,
  session_id uuid not null references sessions(id) on delete cascade,
  status text default 'Not Tested',
  tested_by text default '',
  notes text default '',
  other_observations text default '',
  bug_severity text default '',
  step_checks jsonb default '{}'::jsonb,
  step_notes jsonb default '{}'::jsonb,
  recordings jsonb default '[]'::jsonb,     -- [{url: "..."}]
  screenshots jsonb default '[]'::jsonb,    -- [{path, url, name, size, uploaded_by, uploaded_at}]
  updated_at timestamptz default now(),
  updated_by text default ''
);

-- Idempotent migration for existing deploys that pre-date the screenshots
-- column (the create-table above only fires on fresh installs).
alter table findings add column if not exists screenshots jsonb default '[]'::jsonb;

-- 4. Done marks: per-developer per-test, for the dev portal. Records each
-- developer's fix state and their optional note about the test (shown to the
-- tester, read-only). A row exists if status is set OR note is non-empty.
create table if not exists done_marks (
  test_case_id uuid not null references test_cases(id) on delete cascade,
  session_id uuid not null references sessions(id) on delete cascade,
  developer_name text not null,
  done_at timestamptz default now(),
  status text default 'fixed',              -- 'in_progress' | 'fixed' | null
  note text default '',                     -- dev's note for the tester to read
  primary key (test_case_id, developer_name)
);

-- Indexes
create index if not exists idx_test_cases_session on test_cases(session_id, position);
create index if not exists idx_findings_session on findings(session_id);
create index if not exists idx_done_marks_session on done_marks(session_id);
create index if not exists idx_done_marks_developer on done_marks(session_id, developer_name);

-- Real-time replication: include these tables so the tester portal and dev
-- portal get live push updates when others edit. Wrapped so re-running this
-- file doesn't error on "already member of publication".
do $$
begin
  if not exists (
    select 1 from pg_publication_tables
    where pubname = 'supabase_realtime' and tablename = 'findings'
  ) then
    alter publication supabase_realtime add table findings;
  end if;
  if not exists (
    select 1 from pg_publication_tables
    where pubname = 'supabase_realtime' and tablename = 'done_marks'
  ) then
    alter publication supabase_realtime add table done_marks;
  end if;
end $$;

-- RLS is intentionally NOT enabled for this iteration. The publishable/anon
-- key gives read/write to anyone holding it; access is gated only by the
-- in-app session codes you set per session. If you later host this on a
-- public URL and want tighter access, add RLS policies + Supabase Auth.

-- ─────────────────────────────────────────────────────────────────────────────
-- Storage: screenshots bucket — testers attach evidence images, devs read them.
-- Public bucket so the tester portal, dev portal, and exported reports can
-- render images via plain <img src> without auth handshake. Writes still need
-- explicit policies because RLS on storage.objects is on by default.
-- ─────────────────────────────────────────────────────────────────────────────

insert into storage.buckets (id, name, public)
values ('screenshots', 'screenshots', true)
on conflict (id) do nothing;

-- Permissive insert/delete policies for the anon role (same trust model as
-- the rest of this iteration — gated by in-app session codes, not auth).
do $$
begin
  if not exists (
    select 1 from pg_policies
    where schemaname = 'storage' and tablename = 'objects'
      and policyname = 'screenshots anon insert'
  ) then
    create policy "screenshots anon insert"
    on storage.objects for insert
    to anon
    with check (bucket_id = 'screenshots');
  end if;
  if not exists (
    select 1 from pg_policies
    where schemaname = 'storage' and tablename = 'objects'
      and policyname = 'screenshots anon delete'
  ) then
    create policy "screenshots anon delete"
    on storage.objects for delete
    to anon
    using (bucket_id = 'screenshots');
  end if;
end $$;

-- Bucket-level file size cap (10 MB). Defense in depth — the tester portal
-- already rejects oversized files in JS before upload, but this catches any
-- bypass (direct API call, future client bug, etc.).
update storage.buckets
   set file_size_limit = 10485760
 where id = 'screenshots';
