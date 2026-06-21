-- =====================================================================
-- Gerador de Documentos Infobarra — schema Supabase
-- Cole isto no SQL Editor do projeto novo (ou rode via MCP). Reexecutável.
--
-- ANTES de rodar, em Authentication → Sign In / Up:
--   • DESATIVE "Allow new users to sign up" (você cadastra os usuários no painel).
-- Operadores logam com EMAIL + SENHA (contas criadas em Authentication → Users).
-- =====================================================================

create extension if not exists pgcrypto;

-- ---------- Perfis (1:1 com auth.users) ----------
create table if not exists public.perfis (
  id         uuid primary key references auth.users(id) on delete cascade,
  nome       text not null,
  role       text not null default 'operador' check (role in ('admin','operador')),
  created_at timestamptz not null default now()
);

-- O usuário atual é admin?
create or replace function public.is_admin()
returns boolean language sql security definer stable set search_path = public as $$
  select exists (select 1 from public.perfis where id = auth.uid() and role = 'admin');
$$;

-- Cria o perfil automaticamente ao criar o usuário (o 1º vira admin)
create or replace function public.handle_new_user()
returns trigger language plpgsql security definer set search_path = public as $$
declare v_role text;
begin
  v_role := coalesce(
    new.raw_user_meta_data->>'role',
    case when (select count(*) from public.perfis) = 0 then 'admin' else 'operador' end
  );
  insert into public.perfis (id, nome, role)
  values (new.id, coalesce(new.raw_user_meta_data->>'nome', split_part(new.email,'@',1)), v_role)
  on conflict (id) do nothing;
  return new;
end; $$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users for each row execute function public.handle_new_user();

-- ---------- Clientes (comodato/recibo) ----------
-- Guarda o registro completo em `dados` (jsonb) + colunas úteis para busca.
create table if not exists public.clientes (
  id         uuid primary key default gen_random_uuid(),
  tipo       text not null check (tipo in ('comodato','recibo')),
  contrato   text,
  nome       text,
  cpf        text,
  dados      jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now(),
  updated_by uuid references auth.users(id)
);
create unique index if not exists clientes_tipo_contrato_uk
  on public.clientes (tipo, upper(contrato)) where contrato is not null and contrato <> '';
create index if not exists clientes_tipo_idx on public.clientes(tipo);

-- ---------- Operações / termos (Relatórios) ----------
create table if not exists public.operacoes (
  id          uuid primary key default gen_random_uuid(),
  tipo        text not null check (tipo in ('comodato','recibo')),
  contrato    text,
  cliente     text,
  operador    text,
  operador_id uuid references auth.users(id),
  data        timestamptz not null default now(),
  assinado    boolean not null default false,
  assinado_em timestamptz,
  anexo_url   text
);
create unique index if not exists operacoes_tipo_contrato_uk
  on public.operacoes (tipo, upper(contrato)) where contrato is not null and contrato <> '';

-- ---------- Estoque + catálogo aprendido (compartilhados) ----------
create table if not exists public.estoque (
  id         uuid primary key default gen_random_uuid(),
  marca      text not null,
  modelo     text not null,
  qtd        integer not null default 0,
  updated_at timestamptz not null default now()
);
create unique index if not exists estoque_marca_modelo_uk
  on public.estoque (upper(marca), upper(modelo));

create table if not exists public.catalogo (
  marca   text primary key,
  modelos jsonb not null default '[]'::jsonb
);

-- =====================================================================
-- RLS — equipe interna confiável: qualquer usuário AUTENTICADO lê/grava
-- os dados; a gestão de operadores (perfis) é só de admin.
-- =====================================================================
alter table public.perfis    enable row level security;
alter table public.clientes  enable row level security;
alter table public.operacoes enable row level security;
alter table public.estoque   enable row level security;
alter table public.catalogo  enable row level security;

-- perfis
drop policy if exists perfis_sel on public.perfis;
create policy perfis_sel on public.perfis for select to authenticated using (true);
drop policy if exists perfis_admin_ins on public.perfis;
create policy perfis_admin_ins on public.perfis for insert to authenticated with check (public.is_admin());
drop policy if exists perfis_admin_upd on public.perfis;
create policy perfis_admin_upd on public.perfis for update to authenticated using (public.is_admin()) with check (public.is_admin());
drop policy if exists perfis_admin_del on public.perfis;
create policy perfis_admin_del on public.perfis for delete to authenticated using (public.is_admin() and id <> auth.uid());

-- dados (clientes/operacoes/estoque/catalogo): autenticados podem tudo
do $$
declare t text;
begin
  foreach t in array array['clientes','operacoes','estoque','catalogo'] loop
    execute format('drop policy if exists %1$s_all on public.%1$s;', t);
    execute format('create policy %1$s_all on public.%1$s for all to authenticated using (true) with check (true);', t);
  end loop;
end $$;

-- =====================================================================
-- Storage: bucket dos PDFs assinados (privado, só autenticados)
-- =====================================================================
insert into storage.buckets (id, name, public) values ('assinados','assinados', false)
  on conflict (id) do nothing;
drop policy if exists assinados_rw on storage.objects;
create policy assinados_rw on storage.objects for all to authenticated
  using (bucket_id = 'assinados') with check (bucket_id = 'assinados');
