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
