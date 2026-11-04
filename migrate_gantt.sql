-- =============================================================================
-- migrate_gantt.sql
-- Run this in the Supabase SQL Editor (Dashboard → SQL Editor → New Query)
-- to create the gantt_tasks table that backs the SMTP/TCRS Gantt application.
-- =============================================================================

-- ── 1. Create the table ──────────────────────────────────────────────────────
create table if not exists public.gantt_tasks (
    -- Surrogate PK (auto-incrementing, never exposed to users)
    id            bigint generated always as identity primary key,

    -- Natural key from ClickUp / schedule_data.json
    task_id       text        not null unique,

    -- Hierarchy
    parent_id     text        not null default '',
    parent_task   text        not null default '',
    wbs           text        not null default '',
    level         smallint    not null default 1,

    -- Core task fields  (match schedule_data.json keys exactly)
    project       text        not null default 'SMTP'
                  check (project in ('SMTP', 'TCRS')),
    task_name     text        not null default '',
    task_type     text        not null default 'Task'
                  check (task_type in ('Task', 'Milestone', 'Deliverable', 'Decision Point')),
    status        text        not null default 'To Do'
                  check (status in ('To Do', 'In Progress', 'Complete')),
    start_date    date,
    due_date      date,
    assignees     text        not null default '',
    priority      text        not null default 'Normal'
                  check (priority in ('Urgent', 'High', 'Normal', 'Low')),
    tags          text        not null default '',
    description   text        not null default '',
    task_link     text        not null default '',

    -- Audit columns
    created_at    timestamptz not null default now(),
    updated_at    timestamptz not null default now()
);

-- ── 2. Keep updated_at current automatically ─────────────────────────────────
create or replace function public.set_updated_at()
returns trigger language plpgsql as $$
begin
    new.updated_at = now();
    return new;
end;
$$;

drop trigger if exists gantt_tasks_updated_at on public.gantt_tasks;
create trigger gantt_tasks_updated_at
    before update on public.gantt_tasks
    for each row execute function public.set_updated_at();

-- ── 3. Indexes for the common query patterns ─────────────────────────────────
create index if not exists idx_gantt_tasks_project     on public.gantt_tasks (project);
create index if not exists idx_gantt_tasks_wbs         on public.gantt_tasks (project, wbs);
create index if not exists idx_gantt_tasks_parent_id   on public.gantt_tasks (parent_id);
create index if not exists idx_gantt_tasks_status      on public.gantt_tasks (status);
create index if not exists idx_gantt_tasks_updated_at  on public.gantt_tasks (updated_at);

-- ── 4. Row Level Security ────────────────────────────────────────────────────
-- Enable RLS (all access goes through the policies below)
alter table public.gantt_tasks enable row level security;

-- Allow anonymous reads (the Gantt HTML page reads without login)
create policy "anon_read" on public.gantt_tasks
    for select to anon, authenticated
    using (true);

-- Allow authenticated users (service role, logged-in users) to write
create policy "auth_write" on public.gantt_tasks
    for all to authenticated
    using (true)
    with check (true);

-- ── 5. Comments ───────────────────────────────────────────────────────────────
comment on table  public.gantt_tasks               is 'WSDOT SMTP & TCRS Gantt chart tasks — mirrors schedule_data.json';
comment on column public.gantt_tasks.task_id       is 'ClickUp task ID (natural key, e.g. 86bbh427u)';
comment on column public.gantt_tasks.wbs           is 'Work Breakdown Structure number (e.g. 1.2.3)';
comment on column public.gantt_tasks.level         is 'WBS depth: 1=top-level, 2=child, 3=grandchild, etc.';
comment on column public.gantt_tasks.start_date    is 'Task start date (NULL if not set)';
comment on column public.gantt_tasks.due_date      is 'Task due date (NULL if not set)';

