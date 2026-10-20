create table if not exists public.sheets (
  name text primary key,
  data jsonb not null default '[]'::jsonb,
  ver bigint not null default 1,
  updated_at timestamptz not null default now()
);
alter table public.sheets enable row level security;
revoke all on public.sheets from anon, authenticated;
insert into storage.buckets (id, name, public) values ('files', 'files', false) on conflict (id) do nothing;
update public.sheets set data = (data #>> '{}')::jsonb, ver = ver + 1 where jsonb_typeof(data) = 'string';
select name, ver, jsonb_typeof(data) as t, case when jsonb_typeof(data) = 'array' then jsonb_array_length(data) end as rows from public.sheets order by name;
