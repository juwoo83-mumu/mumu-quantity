
-- 새 Supabase 프로젝트의 SQL Editor에서 실행
create extension if not exists pgcrypto;
create table public.profiles (
 id uuid primary key references auth.users(id) on delete cascade,
 name text not null,
 role text not null default 'employee' check (role in ('employee','owner')),
 approved boolean not null default false,
 created_at timestamptz not null default now()
);
create table public.entries (
 id uuid primary key default gen_random_uuid(),
 employee_id uuid not null references public.profiles(id),
 address text not null,
 duvet integer not null default 0 check(duvet>=0),
 small integer not null default 0 check(small>=0),
 large integer not null default 0 check(large>=0),
 created_at timestamptz not null default now(),
 updated_at timestamptz not null default now()
);
create table public.changes (
 id uuid primary key default gen_random_uuid(),
 entry_id uuid not null references public.entries(id),
 changed_by uuid not null references public.profiles(id),
 reason text not null,
 note text not null default '',
 before_values jsonb not null,
 after_values jsonb not null,
 created_at timestamptz not null default now()
);
create or replace function public.is_owner() returns boolean
language sql stable security definer set search_path='' as $$
 select exists(select 1 from public.profiles where id=auth.uid() and role='owner' and approved);
$$;
create or replace function public.is_approved() returns boolean
language sql stable security definer set search_path='' as $$
 select exists(select 1 from public.profiles where id=auth.uid() and approved);
$$;
create or replace function public.create_profile() returns trigger
language plpgsql security definer set search_path='' as $$
begin
 insert into public.profiles(id,name) values(new.id,coalesce(nullif(new.raw_user_meta_data->>'name',''),split_part(new.email,'@',1)));
 return new;
end; $$;
create trigger after_auth_signup after insert on auth.users
for each row execute function public.create_profile();
alter table public.profiles enable row level security;
alter table public.entries enable row level security;
alter table public.changes enable row level security;
create policy profiles_read on public.profiles for select to authenticated using(id=auth.uid() or public.is_owner());
create policy profiles_owner_update on public.profiles for update to authenticated using(public.is_owner()) with check(public.is_owner());
create policy entries_read on public.entries for select to authenticated using(public.is_owner() or (employee_id=auth.uid() and public.is_approved()));
create policy entries_insert on public.entries for insert to authenticated with check(employee_id=auth.uid() and public.is_approved());
create policy changes_read on public.changes for select to authenticated using(public.is_owner() or (changed_by=auth.uid() and public.is_approved()));
-- 수량 수정 + 이력 기록을 하나의 트랜잭션으로 실행
create or replace function public.revise_entry(p_id uuid,p_address text,p_duvet integer,p_small integer,p_large integer,p_reason text,p_note text)
returns void language plpgsql security definer set search_path='' as $$
declare v public.entries%rowtype;
begin
 if not public.is_approved() then raise exception '승인되지 않은 계정입니다'; end if;
 if p_address is null or btrim(p_address)='' or p_duvet<0 or p_small<0 or p_large<0 or p_reason is null or btrim(p_reason)='' then
 raise exception '수정 값을 확인하세요'; end if;
 select * into v from public.entries where id=p_id for update;
 if not found then raise exception '작업 내역을 찾을 수 없습니다'; end if;
 if v.employee_id<>auth.uid() and not public.is_owner() then raise exception '수정 권한이 없습니다'; end if;
 update public.entries set address=p_address,duvet=p_duvet,small=p_small,large=p_large,updated_at=now() where id=p_id;
 insert into public.changes(entry_id,changed_by,reason,note,before_values,after_values)
 values(p_id,auth.uid(),p_reason,coalesce(p_note,''),
 jsonb_build_object('address',v.address,'duvet',v.duvet,'small',v.small,'large',v.large),
 jsonb_build_object('address',p_address,'duvet',p_duvet,'small',p_small,'large',p_large));
end; $$;
revoke all on function public.revise_entry(uuid,text,integer,integer,integer,text,text) from public;
grant execute on function public.revise_entry(uuid,text,integer,integer,integer,text,text) to authenticated;
-- 사업주 가입 후 본인 이메일로 교체해 1회 실행:
-- update public.profiles p set role='owner',approved=true from auth.users u
-- where p.id=u.id and u.email='OWNER@example.com';
