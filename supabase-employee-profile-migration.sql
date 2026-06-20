-- Employee profile fields used internally by employee management.
-- These fields are not used in payroll calculations.

alter table public.employees
add column if not exists nationality text not null default 'Indian';

alter table public.employees
add column if not exists state text not null default '';

update public.employees
set nationality = 'Indian'
where nationality is null or btrim(nationality) = '';
