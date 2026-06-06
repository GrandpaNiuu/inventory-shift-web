-- Supabase PostgreSQL schema for the production version.
-- The current MVP uses browser localStorage so it can be tested immediately after Vercel deployment.

create table if not exists public.profiles (
  id uuid primary key,
  email text unique,
  display_name text,
  role text not null default 'staff' check (role in ('owner', 'admin', 'morning_staff', 'evening_staff', 'staff')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.products (
  id uuid primary key default gen_random_uuid(),
  sku text not null unique,
  name text not null,
  category text,
  spec text,
  system_stock numeric(12, 2) not null default 0,
  warning_stock numeric(12, 2) not null default 0,
  is_active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.shift_income (
  id uuid primary key default gen_random_uuid(),
  date date not null,
  shift text not null check (shift in ('morning', 'evening')),
  cash numeric(12, 2) not null default 0,
  wechat numeric(12, 2) not null default 0,
  alipay numeric(12, 2) not null default 0,
  platform numeric(12, 2) not null default 0,
  refund numeric(12, 2) not null default 0,
  expense numeric(12, 2) not null default 0,
  expected_total numeric(12, 2) generated always as (cash + wechat + alipay + platform - refund - expense) stored,
  actual_total numeric(12, 2) not null default 0,
  note text,
  operator_id uuid,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique(date, shift)
);

create table if not exists public.stock_count_sessions (
  id uuid primary key default gen_random_uuid(),
  week_start date not null,
  count_date date not null default current_date,
  status text not null default 'draft' check (status in ('draft', 'confirmed')),
  operator_id uuid,
  note text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.stock_count_items (
  id uuid primary key default gen_random_uuid(),
  session_id uuid not null references public.stock_count_sessions(id) on delete cascade,
  sku text not null,
  product_id uuid references public.products(id),
  system_stock numeric(12, 2) not null default 0,
  counted_stock numeric(12, 2),
  difference numeric(12, 2) generated always as (counted_stock - system_stock) stored,
  note text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique(session_id, sku)
);

create table if not exists public.audit_logs (
  id uuid primary key default gen_random_uuid(),
  user_id uuid,
  action text not null,
  table_name text not null,
  record_id text,
  before_data jsonb,
  after_data jsonb,
  created_at timestamptz not null default now()
);

create index if not exists idx_shift_income_date on public.shift_income(date);
create index if not exists idx_shift_income_month on public.shift_income(date, shift);
create index if not exists idx_stock_count_week on public.stock_count_sessions(week_start);
create index if not exists idx_stock_count_items_session on public.stock_count_items(session_id);
