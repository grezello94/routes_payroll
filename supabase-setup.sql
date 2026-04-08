create table if not exists public.users (
  id text primary key,
  username text not null,
  username_lc text not null unique,
  email text not null,
  email_lc text not null unique,
  email_verified boolean not null default false,
  email_verification_hash text not null default '',
  email_verification_expires_at text not null default '',
  password_hash text not null default '',
  reset_code_hash text not null default '',
  reset_code_expires_at text not null default '',
  created_at timestamptz not null default now()
);

create table if not exists public.companies (
  id bigint primary key,
  name text not null unique,
  name_lc text not null unique,
  logo_data_url text not null default '',
  created_at timestamptz not null default now()
);

create table if not exists public.designation_presets (
  id bigint primary key,
  company_id bigint not null references public.companies(id) on delete cascade,
  name text not null,
  name_lc text not null,
  position_index integer not null default 0,
  created_at timestamptz not null default now(),
  unique (company_id, name_lc)
);

create table if not exists public.employees (
  id bigint primary key,
  company_id bigint not null references public.companies(id) on delete cascade,
  employee_id text not null,
  employee_name text not null,
  joining_date text not null default '',
  birth_date text not null default '',
  base_salary numeric not null default 0,
  opening_advance numeric not null default 0,
  designation text not null default '',
  mobile_number text not null default '',
  status text not null default 'working',
  leave_from text not null default '',
  leave_to text not null default '',
  terminated_on text not null default '',
  notes text not null default '',
  position_index integer not null default 0,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (company_id, employee_id)
);

create table if not exists public.payroll_entries (
  id bigint primary key,
  company_id bigint not null references public.companies(id) on delete cascade,
  month text not null,
  employee_id text not null,
  employee_name text not null,
  designation text not null,
  present_salary numeric not null default 0,
  increment numeric not null default 0,
  old_advance_taken numeric not null default 0,
  extra_advance_added numeric not null default 0,
  deduction_entered numeric not null default 0,
  days_absent numeric not null default 0,
  comment text not null default '',
  position_index integer not null default 0,
  updated_at timestamptz not null default now()
);

create table if not exists public.payroll_reports (
  id bigint primary key,
  company_id bigint not null references public.companies(id) on delete cascade,
  month text not null,
  checked_at text not null default '',
  generated_at text not null default '',
  employee_count integer not null default 0,
  snapshot_json text not null default '',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (company_id, month)
);

create index if not exists idx_designation_company_position on public.designation_presets(company_id, position_index);
create index if not exists idx_employees_company_position on public.employees(company_id, position_index);
create index if not exists idx_payroll_company_month_position on public.payroll_entries(company_id, month, position_index);
create index if not exists idx_payroll_reports_company_month on public.payroll_reports(company_id, month);

insert into public.companies (id, name, name_lc, logo_data_url)
values (1, 'Routes Payroll', 'routes payroll', '')
on conflict (id) do nothing;
