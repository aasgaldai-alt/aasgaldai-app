create table if not exists public.sheets (
  name text primary key,
  data jsonb not null default '[]'::jsonb,
  ver bigint not null default 1,
  updated_at timestamptz not null default now()
);
alter table public.sheets enable row level security;
revoke all on public.sheets from anon, authenticated;
insert into storage.buckets (id, name, public) values ('files', 'files', false) on conflict (id) do nothing;
select name, ver, jsonb_array_length(data) as rows from public.sheets order by name;
