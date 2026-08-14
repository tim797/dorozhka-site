-- Выполните этот SQL в Supabase → SQL Editor.
create extension if not exists pgcrypto;

create table if not exists public.meets (
  id uuid primary key default gen_random_uuid(),
  title text not null,
  date_label text,
  pool text,
  records jsonb not null default '[]'::jsonb,
  uploaded_at timestamptz not null default now()
);

alter table public.meets enable row level security;

-- Публичное чтение результатов.
create policy "public read meets"
on public.meets for select
to anon, authenticated
using (true);

-- Для первого теста разрешаем публичную загрузку и удаление.
-- После проверки лучше заменить эти политики авторизацией тренера/админа.
create policy "public insert meets"
on public.meets for insert
to anon, authenticated
with check (true);

create policy "public delete meets"
on public.meets for delete
to anon, authenticated
using (true);

create index if not exists meets_uploaded_at_idx on public.meets(uploaded_at desc);