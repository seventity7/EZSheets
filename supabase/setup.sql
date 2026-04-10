create extension if not exists pgcrypto;

create table if not exists public.sheets (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references auth.users(id) on delete cascade,
  title text not null,
  code text not null unique,
  rows_count integer not null default 30 check (rows_count between 1 and 200),
  cols_count integer not null default 12 check (cols_count between 1 and 50),
  default_role text not null default 'viewer' check (default_role in ('viewer', 'editor')),
  data jsonb not null default '{"cells": {}}'::jsonb,
  version bigint not null default 1,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.sheet_members (
  sheet_id uuid not null references public.sheets(id) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  role text not null check (role in ('viewer', 'editor')),
  created_at timestamptz not null default now(),
  primary key (sheet_id, user_id)
);

create or replace function public.touch_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_sheets_touch_updated_at on public.sheets;
create trigger trg_sheets_touch_updated_at
before update on public.sheets
for each row execute procedure public.touch_updated_at();

alter table public.sheets enable row level security;
alter table public.sheet_members enable row level security;

grant usage on schema public to anon, authenticated;
grant select, insert, update, delete on public.sheets to authenticated;
grant select on public.sheet_members to authenticated;

create or replace function public.join_sheet_by_code(p_code text)
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  v_sheet_id uuid;
  v_owner_id uuid;
  v_default_role text;
begin
  if auth.uid() is null then
    raise exception 'You must be authenticated to join a sheet.';
  end if;

  select s.id, s.owner_id, s.default_role
    into v_sheet_id, v_owner_id, v_default_role
  from public.sheets s
  where s.code = upper(trim(p_code));

  if v_sheet_id is null then
    raise exception 'No sheet exists for that code.';
  end if;

  if auth.uid() = v_owner_id then
    return v_sheet_id;
  end if;

  insert into public.sheet_members (sheet_id, user_id, role)
  values (v_sheet_id, auth.uid(), v_default_role)
  on conflict (sheet_id, user_id) do update
    set role = excluded.role;

  return v_sheet_id;
end;
$$;

grant execute on function public.join_sheet_by_code(text) to authenticated;

drop policy if exists "owners can insert sheets" on public.sheets;
drop policy if exists "owners and members can read accessible sheets" on public.sheets;
drop policy if exists "owners and editors can update sheets" on public.sheets;
drop policy if exists "owners can delete sheets" on public.sheets;
drop policy if exists "users can read their memberships" on public.sheet_members;

create policy "owners can insert sheets"
on public.sheets
for insert
with check (auth.uid() = owner_id);

create policy "owners and members can read accessible sheets"
on public.sheets
for select
using (
  auth.uid() = owner_id
  or exists (
    select 1
    from public.sheet_members sm
    where sm.sheet_id = sheets.id
      and sm.user_id = auth.uid()
  )
);

create policy "owners and editors can update sheets"
on public.sheets
for update
using (
  auth.uid() = owner_id
  or exists (
    select 1
    from public.sheet_members sm
    where sm.sheet_id = sheets.id
      and sm.user_id = auth.uid()
      and sm.role = 'editor'
  )
)
with check (
  auth.uid() = owner_id
  or exists (
    select 1
    from public.sheet_members sm
    where sm.sheet_id = sheets.id
      and sm.user_id = auth.uid()
      and sm.role = 'editor'
  )
);

create policy "owners can delete sheets"
on public.sheets
for delete
using (auth.uid() = owner_id);

create policy "users can read their memberships"
on public.sheet_members
for select
using (
  auth.uid() = user_id
  or exists (
    select 1
    from public.sheets s
    where s.id = sheet_members.sheet_id
      and s.owner_id = auth.uid()
  )
);

create or replace function public.user_has_sheet_permission(p_sheet_id uuid, p_user_id uuid, p_permission text)
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  with sheet_row as (
    select owner_id, data
    from public.sheets
    where id = p_sheet_id
  ), member_row as (
    select profile
    from sheet_row s,
         lateral jsonb_array_elements(coalesce(s.data->'settings'->'memberProfiles', '[]'::jsonb)) as profile
    where coalesce(profile->>'userId', '') = p_user_id::text
    limit 1
  )
  select exists (
           select 1
           from sheet_row s
           where s.owner_id = p_user_id)
         or exists (
           select 1
           from member_row m
           where coalesce((m.profile->>'isBlocked')::boolean, false) = false
             and (
               (lower(p_permission) <> 'deletesheet' and coalesce((m.profile->'permissions'->>'admin')::boolean, false))
               or case lower(p_permission)
                    when 'editsheet' then coalesce((m.profile->'permissions'->>'editSheet')::boolean, false)
                    when 'deletesheet' then coalesce((m.profile->'permissions'->>'deleteSheet')::boolean, false)
                    when 'editpermissions' then coalesce((m.profile->'permissions'->>'editPermissions')::boolean, false)
                    when 'createtabs' then coalesce((m.profile->'permissions'->>'createTabs')::boolean, false)
                    when 'seehistory' then coalesce((m.profile->'permissions'->>'seeHistory')::boolean, false)
                    when 'usecomments' then coalesce((m.profile->'permissions'->>'useComments')::boolean, false)
                    when 'importsheet' then coalesce((m.profile->'permissions'->>'importSheet')::boolean, false)
                    when 'savelocal' then coalesce((m.profile->'permissions'->>'saveLocal')::boolean, false)
                    when 'invite' then coalesce((m.profile->'permissions'->>'invite')::boolean, false)
                    when 'blockusers' then coalesce((m.profile->'permissions'->>'blockUsers')::boolean, false)
                    when 'admin' then coalesce((m.profile->'permissions'->>'admin')::boolean, false)
                    else false
                  end
             )
         );
$$;

grant execute on function public.user_has_sheet_permission(uuid, uuid, text) to authenticated;

create or replace function public.save_sheet_if_permitted(
  p_sheet_id uuid,
  p_expected_version bigint,
  p_title text,
  p_rows_count integer,
  p_cols_count integer,
  p_default_role text,
  p_data jsonb)
returns setof public.sheets
language plpgsql
security definer
set search_path = public
as $$
declare
  v_sheet public.sheets%rowtype;
  v_uid uuid;
  v_can_save boolean;
begin
  v_uid := auth.uid();
  if v_uid is null then
    raise exception 'You must be authenticated to update a sheet.';
  end if;

  select *
    into v_sheet
  from public.sheets
  where id = p_sheet_id
  for update;

  if not found then
    raise exception 'Sheet not found.';
  end if;

  if v_sheet.version <> p_expected_version then
    raise exception 'version conflict';
  end if;

  v_can_save := v_sheet.owner_id = v_uid
    or exists (
      select 1
      from public.sheet_members sm
      where sm.sheet_id = p_sheet_id
        and sm.user_id = v_uid
        and sm.role = 'editor'
    )
    or public.user_has_sheet_permission(p_sheet_id, v_uid, 'editsheet')
    or public.user_has_sheet_permission(p_sheet_id, v_uid, 'editpermissions')
    or public.user_has_sheet_permission(p_sheet_id, v_uid, 'createtabs')
    or public.user_has_sheet_permission(p_sheet_id, v_uid, 'usecomments')
    or public.user_has_sheet_permission(p_sheet_id, v_uid, 'seehistory')
    or public.user_has_sheet_permission(p_sheet_id, v_uid, 'importsheet')
    or public.user_has_sheet_permission(p_sheet_id, v_uid, 'admin');

  if not v_can_save then
    raise exception 'You do not have permission to update this sheet.';
  end if;

  update public.sheets
     set title = coalesce(nullif(trim(p_title), ''), v_sheet.title),
         rows_count = greatest(1, least(200, p_rows_count)),
         cols_count = greatest(1, least(50, p_cols_count)),
         default_role = case when lower(coalesce(p_default_role, 'viewer')) = 'editor' then 'editor' else 'viewer' end,
         data = coalesce(p_data, v_sheet.data),
         version = v_sheet.version + 1,
         updated_at = now()
   where id = p_sheet_id;

  return query
  select *
  from public.sheets
  where id = p_sheet_id;
end;
$$;

grant execute on function public.save_sheet_if_permitted(uuid, bigint, text, integer, integer, text, jsonb) to authenticated;

create or replace function public.delete_sheet_if_permitted(p_sheet_id uuid)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  v_uid uuid;
  v_owner_id uuid;
begin
  v_uid := auth.uid();
  if v_uid is null then
    raise exception 'You must be authenticated to delete a sheet.';
  end if;

  select owner_id
    into v_owner_id
  from public.sheets
  where id = p_sheet_id;

  if v_owner_id is null then
    raise exception 'Sheet not found.';
  end if;

  if v_owner_id <> v_uid and not public.user_has_sheet_permission(p_sheet_id, v_uid, 'deletesheet') then
    raise exception 'You do not have permission to delete this sheet.';
  end if;

  delete from public.sheets where id = p_sheet_id;
end;
$$;

grant execute on function public.delete_sheet_if_permitted(uuid) to authenticated;

create or replace function public.list_sheet_members_if_permitted(p_sheet_id uuid)
returns table(sheet_id uuid, user_id uuid, role text, created_at timestamptz)
language plpgsql
security definer
set search_path = public
as $$
declare
  v_uid uuid;
  v_owner_id uuid;
begin
  v_uid := auth.uid();
  if v_uid is null then
    raise exception 'You must be authenticated to view sheet members.';
  end if;

  select owner_id into v_owner_id from public.sheets where id = p_sheet_id;
  if v_owner_id is null then
    raise exception 'Sheet not found.';
  end if;

  if v_owner_id <> v_uid
     and not public.user_has_sheet_permission(p_sheet_id, v_uid, 'editpermissions')
     and not public.user_has_sheet_permission(p_sheet_id, v_uid, 'blockusers')
     and not public.user_has_sheet_permission(p_sheet_id, v_uid, 'admin') then
    raise exception 'You do not have permission to view sheet members.';
  end if;

  return query
  select sm.sheet_id, sm.user_id, sm.role, sm.created_at
  from public.sheet_members sm
  where sm.sheet_id = p_sheet_id
  order by sm.created_at asc;
end;
$$;

grant execute on function public.list_sheet_members_if_permitted(uuid) to authenticated;

create or replace function public.remove_sheet_member_if_permitted(p_sheet_id uuid, p_user_id uuid)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  v_uid uuid;
  v_owner_id uuid;
begin
  v_uid := auth.uid();
  if v_uid is null then
    raise exception 'You must be authenticated to remove a sheet member.';
  end if;

  select owner_id into v_owner_id from public.sheets where id = p_sheet_id;
  if v_owner_id is null then
    raise exception 'Sheet not found.';
  end if;

  if p_user_id = v_owner_id then
    raise exception 'The sheet owner cannot be removed.';
  end if;

  if v_owner_id <> v_uid
     and not public.user_has_sheet_permission(p_sheet_id, v_uid, 'blockusers')
     and not public.user_has_sheet_permission(p_sheet_id, v_uid, 'admin') then
    raise exception 'You do not have permission to remove a sheet member.';
  end if;

  delete from public.sheet_members
  where sheet_id = p_sheet_id
    and user_id = p_user_id;
end;
$$;

grant execute on function public.remove_sheet_member_if_permitted(uuid, uuid) to authenticated;

-- EZSheets live collaboration tables
create table if not exists public.sheetsync_sheet_presence (
  sheet_id uuid not null references public.sheetsync_sheets(id) on delete cascade,
  user_id text not null,
  user_name text not null default '',
  active_tab_name text null,
  editing_cell_key text null,
  last_seen_utc timestamptz not null default now(),
  primary key (sheet_id, user_id)
);

create index if not exists idx_sheetsync_sheet_presence_sheet_seen
  on public.sheetsync_sheet_presence (sheet_id, last_seen_utc desc);

create table if not exists public.sheetsync_sheet_chat_messages (
  id uuid primary key default gen_random_uuid(),
  sheet_id uuid not null references public.sheetsync_sheets(id) on delete cascade,
  author_user_id text not null,
  author_name text not null default '',
  message text not null,
  created_at timestamptz not null default now()
);

create index if not exists idx_sheetsync_sheet_chat_messages_sheet_created
  on public.sheetsync_sheet_chat_messages (sheet_id, created_at asc);

create table if not exists public.sheetsync_sheet_cell_locks (
  sheet_id uuid not null references public.sheetsync_sheets(id) on delete cascade,
  cell_key text not null,
  user_id text not null,
  user_name text not null default '',
  locked_at timestamptz not null default now(),
  expires_at timestamptz not null default now() + interval '15 seconds',
  primary key (sheet_id, cell_key)
);

create index if not exists idx_sheetsync_sheet_cell_locks_sheet_expires
  on public.sheetsync_sheet_cell_locks (sheet_id, expires_at desc);

create table if not exists public.sheetsync_sheet_unique_codes (
  id uuid primary key default gen_random_uuid(),
  sheet_id uuid not null references public.sheetsync_sheets(id) on delete cascade,
  code text not null unique,
  created_by_user_id text not null,
  created_at timestamptz not null default now(),
  used_at timestamptz null,
  used_by_user_id text null,
  invalidated_at timestamptz null,
  invalidated_by_user_id text null
);

create index if not exists idx_sheetsync_sheet_unique_codes_sheet_active
  on public.sheetsync_sheet_unique_codes (sheet_id, created_at desc)
  where used_at is null and invalidated_at is null;

create table if not exists public.sheetsync_sheet_blocklist (
  sheet_id uuid not null references public.sheetsync_sheets(id) on delete cascade,
  user_id text not null,
  character_name text not null default '',
  reason text not null default '',
  removed_at timestamptz not null default now(),
  removed_by_user_id text not null,
  primary key (sheet_id, user_id)
);

create index if not exists idx_sheetsync_sheet_blocklist_sheet_removed
  on public.sheetsync_sheet_blocklist (sheet_id, removed_at desc);


create table if not exists public.sheetsync_discord_sessions (
  token_hash text primary key,
  user_id text not null,
  user_display_name text not null default '',
  user_email text not null default '',
  expires_at timestamptz not null,
  created_at timestamptz not null default now(),
  last_seen_at timestamptz not null default now()
);

create index if not exists idx_sheetsync_discord_sessions_user_id
  on public.sheetsync_discord_sessions (user_id);

create index if not exists idx_sheetsync_discord_sessions_expires_at
  on public.sheetsync_discord_sessions (expires_at);
