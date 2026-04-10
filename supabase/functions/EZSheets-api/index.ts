import { createClient } from 'npm:@supabase/supabase-js@2'

type DiscordTokenResponse = {
  access_token: string
  refresh_token?: string
  expires_in: number
  token_type?: string
  scope?: string
}

type DiscordUser = {
  id: string
  username?: string
  global_name?: string | null
  avatar?: string | null
  email?: string | null
}

type SheetRow = {
  id: string
  owner_id: string
  title: string
  code: string
  rows_count: number
  cols_count: number
  default_role: string
  data: any
  version: number
  created_at: string
  updated_at: string
}

type SheetMetaRow = {
  id: string
  owner_id: string
  title: string
  version: number
  updated_at: string
}

const supabaseUrl = Deno.env.get('SUPABASE_URL')
const serviceRoleKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')
const discordClientId = Deno.env.get('DISCORD_CLIENT_ID')
const discordClientSecret = Deno.env.get('DISCORD_CLIENT_SECRET')

if (!supabaseUrl || !serviceRoleKey) {
  throw new Error('Missing SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY.')
}

const admin = createClient(supabaseUrl, serviceRoleKey)

type CachedServerSession = {
  user: DiscordUser
  expiresAtMs: number
}

const serverSessionCache = new Map<string, CachedServerSession>()

const jsonHeaders = {
  'Content-Type': 'application/json; charset=utf-8',
}

Deno.serve(async (req: Request) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', {
      headers: {
        ...jsonHeaders,
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'authorization, apikey, content-type, x-EZSheets-session, x-sheetsync-session',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
      },
    })
  }

  if (req.method !== 'POST') {
    return json({ message: 'Method not allowed.' }, 405)
  }

  let body: any
  try {
    body = await req.json()
  } catch {
    return json({ message: 'Invalid JSON body.' }, 400)
  }

  const action = `${body?.action ?? ''}`.trim()
  const payload = body?.payload ?? {}

  try {
    switch (action) {
      case 'exchange_code':
        return await handleExchangeCode(payload)
      case 'refresh_session':
        return await handleRefreshSession(payload)
      case 'logout':
        return await handleLogout(payload)
      case 'list_sheets':
        return await handleListSheets(req)
      case 'get_sheet':
        return await handleGetSheet(req, payload)
      case 'get_access_role':
        return await handleGetAccessRole(req, payload)
      case 'create_sheet':
        return await handleCreateSheet(req, payload)
      case 'update_sheet':
        return await handleUpdateSheet(req, payload)
      case 'join_sheet_by_code':
        return await handleJoinSheetByCode(req, payload)
      case 'delete_sheet':
        return await handleDeleteSheet(req, payload)
      case 'list_sheet_members':
        return await handleListSheetMembers(req, payload)
      case 'remove_sheet_member':
        return await handleRemoveSheetMember(req, payload)
      case 'list_sheet_blocklist':
        return await handleListSheetBlocklist(req, payload)
      case 'unblock_sheet_member':
        return await handleUnblockSheetMember(req, payload)
      case 'sync_presence':
        return await handleSyncPresence(req, payload)
      case 'post_chat_message':
        return await handlePostChatMessage(req, payload)
      case 'get_sheet_runtime':
        return await handleGetSheetRuntime(req, payload)
      case 'acquire_cell_lock':
        return await handleAcquireCellLock(req, payload)
      case 'release_cell_lock':
        return await handleReleaseCellLock(req, payload)
      case 'create_unique_code':
        return await handleCreateUniqueCode(req, payload, false)
      case 'invalidate_and_create_unique_code':
        return await handleCreateUniqueCode(req, payload, true)
      default:
        return json({ message: 'Unknown action.' }, 400)
    }
  } catch (error) {
    console.error('EZSheets-api error', error)
    const message = error instanceof Error ? error.message : 'Unexpected server error.'
    const status = error instanceof HttpError ? error.status : 500
    return json({ message }, status)
  }
})

async function handleExchangeCode(payload: any): Promise<Response> {
  const code = `${payload?.code ?? ''}`.trim()
  const redirectUri = `${payload?.redirect_uri ?? ''}`.trim()
  if (!code || !redirectUri) {
    return json({ message: 'Missing code or redirect_uri.' }, 400)
  }

  const token = await exchangeDiscordToken({
    grant_type: 'authorization_code',
    code,
    redirect_uri: redirectUri,
  })

  const user = await fetchDiscordUser(token.access_token)
  return json(await buildSession(token, user))
}

async function handleRefreshSession(payload: any): Promise<Response> {
  const refreshToken = `${payload?.refresh_token ?? ''}`.trim()
  if (!refreshToken) {
    return json({ message: 'Missing refresh_token.' }, 400)
  }

  const previousServerSessionToken = `${payload?.server_session_token ?? ''}`.trim()

  const token = await exchangeDiscordToken({
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
  })

  const user = await fetchDiscordUser(token.access_token)
  return json(await buildSession(token, user, previousServerSessionToken))
}

async function handleLogout(payload: any): Promise<Response> {
  const refreshToken = `${payload?.refresh_token ?? ''}`.trim()
  const accessToken = `${payload?.access_token ?? ''}`.trim()
  const serverSessionToken = `${payload?.server_session_token ?? ''}`.trim()

  if (serverSessionToken) {
    await revokeServerSession(serverSessionToken)
  }

  if (refreshToken) {
    await revokeDiscordToken(refreshToken)
  } else if (accessToken) {
    await revokeDiscordToken(accessToken)
  }

  return json({ success: true })
}

async function handleListSheets(req: Request): Promise<Response> {
  const user = await requireDiscordUser(req)

  const { data: ownedSheets, error: ownedError } = await admin
    .from('sheetsync_sheets')
    .select('id, owner_id, title, code, rows_count, cols_count, default_role, version, created_at, updated_at')
    .eq('owner_id', user.id)
    .order('created_at', { ascending: false })

  if (ownedError) {
    return dbError(ownedError)
  }

  const { data: memberships, error: membershipError } = await admin
    .from('sheetsync_sheet_members')
    .select('sheet_id, role')
    .eq('user_id', user.id)

  if (membershipError) {
    return dbError(membershipError)
  }

  const membershipBySheetId = new Map<string, any>()
  for (const row of memberships ?? []) {
    const key = `${row?.sheet_id ?? ''}`
    if (key) {
      membershipBySheetId.set(key, row)
    }
  }

  const results = new Map<string, any>()
  for (const row of ownedSheets ?? []) {
    results.set(row.id, {
      ...row,
      user_role: 'Owner',
      user_role_color: 0xffd99a32,
    })
  }

  const missingIds = Array.from(membershipBySheetId.keys()).filter((id) => !results.has(id))
  if (missingIds.length > 0) {
    const { data: memberSheets, error: memberSheetsError } = await admin
      .from('sheetsync_sheets')
      .select('id, owner_id, title, code, rows_count, cols_count, default_role, version, created_at, updated_at')
      .in('id', missingIds)

    if (memberSheetsError) {
      return dbError(memberSheetsError)
    }

    for (const row of memberSheets ?? []) {
      const membership = membershipBySheetId.get(`${row.id}`)
      const userRole = normalizeRole(membership?.role ?? row.default_role)
      results.set(row.id, {
        ...row,
        user_role: userRole === 'editor' ? 'Editor' : 'Viewer',
        user_role_color: 0xff5c5c5c,
      })
    }
  }

  const ordered = Array.from(results.values())
    .sort((a, b) => (Date.parse(`${b.updated_at ?? ''}`) - Date.parse(`${a.updated_at ?? ''}`)) || `${b.id ?? ''}`.localeCompare(`${a.id ?? ''}`))
  return json(ordered)
}

async function handleGetSheet(req: Request, payload: any): Promise<Response> {
  const user = await requireDiscordUser(req)
  const sheetId = `${payload?.sheet_id ?? ''}`.trim()
  if (!sheetId) {
    return json({ message: 'Missing sheet_id.' }, 400)
  }

  const sheet = await getSheetOrResponse(sheetId)
  if (sheet instanceof Response) {
    return sheet
  }

  const role = await getRoleForUser(sheet, user.id)
  if (!role || isUserBlocked(sheet, user.id)) {
    return json({ message: 'You do not have access to this sheet.' }, 403)
  }

  return json(sheet)
}

async function handleGetAccessRole(req: Request, payload: any): Promise<Response> {
  const user = await requireDiscordUser(req)
  const sheetId = `${payload?.sheet_id ?? ''}`.trim()
  if (!sheetId) {
    return json({ message: 'Missing sheet_id.' }, 400)
  }

  const sheet = await getSheetOrResponse(sheetId)
  if (sheet instanceof Response) {
    return sheet
  }

  const role = await getRoleForUser(sheet, user.id)
  if (!role || isUserBlocked(sheet, user.id)) {
    return json({ message: 'You do not have access to this sheet.' }, 403)
  }

  return json({ role })
}

async function handleCreateSheet(req: Request, payload: any): Promise<Response> {
  const user = await requireDiscordUser(req)
  const characterName = `${payload?.character_name ?? user.global_name ?? user.username ?? 'Discord user'}`.trim()
  const normalizedData = ensureSheetData(payload?.data ?? { tabs: [], activeTabIndex: 0, settings: {} })
  ensureOwnerProfile(normalizedData, user.id)
  const ownerProfile = getOrCreateProfile(normalizedData, user.id)
  ownerProfile.characterName = characterName
  ownerProfile.assignedPresetName = 'Owner'
  ownerProfile.roleColor = 0xffd99a32
  ownerProfile.isBlocked = false
  ownerProfile.joinedAtUtc = ownerProfile.joinedAtUtc || new Date().toISOString()
  ownerProfile.lastSeenUtc = new Date().toISOString()

  const insertRow = {
    owner_id: user.id,
    title: `${payload?.title ?? 'Untitled sheet'}`.trim() || 'Untitled sheet',
    code: `${payload?.code ?? ''}`.trim().toUpperCase(),
    rows_count: Number(payload?.rows_count ?? 30),
    cols_count: Number(payload?.cols_count ?? 12),
    default_role: normalizeDefaultRoleValue(payload?.default_role),
    data: normalizedData,
    version: Number(payload?.version ?? 1),
  }

  const { data, error } = await admin
    .from('sheetsync_sheets')
    .insert(insertRow)
    .select('id, owner_id, title, code, rows_count, cols_count, default_role, version, created_at, updated_at, data')
    .single()

  if (error) {
    if (`${error.message}`.toLowerCase().includes('duplicate') || error.code === '23505') {
      return json({ message: 'Share code already exists.' }, 409)
    }

    return dbError(error)
  }

  try {
    await admin.from('sheetsync_sheet_unique_codes').insert({
      sheet_id: data.id,
      code: generateJoinCode(12),
      created_by_user_id: user.id,
    })
  } catch {
  }

  return json(data)
}

async function handleUpdateSheet(req: Request, payload: any): Promise<Response> {
  const user = await requireDiscordUser(req)
  const sheetId = `${payload?.id ?? ''}`.trim()
  if (!sheetId) {
    return json({ message: 'Missing sheet id.' }, 400)
  }

  let sheetOrResponse = await getSheetOrResponse(sheetId)
  if (sheetOrResponse instanceof Response) {
    return sheetOrResponse
  }

  let sheet = sheetOrResponse
  const role = await getRoleForUser(sheet, user.id)
  const canEdit = role === 'editor'
    || sheet.owner_id === user.id
    || hasSheetPermissionInData(sheet, user.id, 'editSheet')
    || hasSheetPermissionInData(sheet, user.id, 'editPermissions')
    || hasSheetPermissionInData(sheet, user.id, 'createTabs')
    || hasSheetPermissionInData(sheet, user.id, 'useComments')
    || hasSheetPermissionInData(sheet, user.id, 'seeHistory')
    || hasSheetPermissionInData(sheet, user.id, 'importSheet')
    || hasSheetPermissionInData(sheet, user.id, 'admin')

  if (!canEdit || isUserBlocked(sheet, user.id)) {
    return json({ message: 'You do not have permission to edit this sheet.' }, 403)
  }

  for (let attempt = 0; attempt < 5; attempt++) {
    const mergedData = mergeSheetData(sheet.data, payload?.data, sheet.owner_id)
    const targetVersion = Number(sheet.version ?? 0) + 1
    const { data, error } = await admin
      .from('sheetsync_sheets')
      .update({
        title: `${payload?.title ?? sheet.title}`.trim() || sheet.title,
        rows_count: Number(payload?.rows_count ?? sheet.rows_count),
        cols_count: Number(payload?.cols_count ?? sheet.cols_count),
        default_role: normalizeDefaultRoleValue(payload?.default_role ?? sheet.default_role),
        data: mergedData,
        version: targetVersion,
      })
      .eq('id', sheetId)
      .eq('version', sheet.version)
      .select('id, owner_id, title, code, rows_count, cols_count, default_role, version, created_at, updated_at, data')
      .maybeSingle()

    if (error) {
      return dbError(error)
    }

    if (data) {
      return json(data)
    }

    const latest = await getSheetOrResponse(sheetId)
    if (latest instanceof Response) {
      return latest
    }
    sheet = latest
  }

  return json({ message: 'This sheet changed on the server before your save completed. Download the latest version and try again.' }, 409)
}

async function handleJoinSheetByCode(req: Request, payload: any): Promise<Response> {
  const user = await requireDiscordUser(req)
  const code = `${payload?.code ?? ''}`.trim().toUpperCase()
  if (!code) {
    return json({ message: 'Missing sheet code.' }, 400)
  }

  const resolved = await resolveSheetByJoinCode(code)
  if (resolved instanceof Response) {
    return resolved
  }

  const { sheet, uniqueCodeRow, temporaryInviteCode, codeType } = resolved
  if (await isUserBlockedFromSheet(sheet.id, user.id) || isUserBlocked(sheet, user.id)) {
    return json({ message: 'You are blocked from this sheet.' }, 403)
  }

  const joinedAtIso = new Date().toISOString()
  const memberRole = deriveMembershipRoleFromDefaultRole(sheet.default_role, sheet.data)
  const existingMember = await admin
    .from('sheetsync_sheet_members')
    .select('user_id')
    .eq('sheet_id', sheet.id)
    .eq('user_id', user.id)
    .maybeSingle()

  if (existingMember.error) {
    return dbError(existingMember.error)
  }

  const isFirstMembershipJoin = !existingMember.data?.user_id && sheet.owner_id !== user.id

  if (sheet.owner_id !== user.id) {
    const { error: upsertError } = await admin
      .from('sheetsync_sheet_members')
      .upsert(
        {
          sheet_id: sheet.id,
          user_id: user.id,
          role: memberRole,
        },
        { onConflict: 'sheet_id,user_id' },
      )

    if (upsertError) {
      return dbError(upsertError)
    }
  }

  if (uniqueCodeRow) {
    const { error: consumeError } = await admin
      .from('sheetsync_sheet_unique_codes')
      .update({ used_at: joinedAtIso, used_by_user_id: user.id })
      .eq('id', uniqueCodeRow.id)
      .is('used_at', null)
      .is('invalidated_at', null)

    if (consumeError) {
      return dbError(consumeError)
    }
  }

  const characterName = `${payload?.character_name ?? user.global_name ?? user.username ?? 'Discord user'}`.trim()
  await mutateSheetData(sheet.id, (data: any) => {
    const profile = getOrCreateProfile(data, user.id)
    profile.characterName = characterName
    profile.joinedAtUtc = profile.joinedAtUtc || joinedAtIso
    profile.lastSeenUtc = joinedAtIso
    profile.isBlocked = false
    profile.accessExpiresAtUtc = null
    if (!profile.assignedPresetName) {
      profile.assignedPresetName = 'Viewer'
    }
    if (!profile.roleColor) {
      profile.roleColor = 0xff5c5c5c
    }
    applyDefaultRoleToProfile(profile, data, sheet.default_role)

    if (codeType === 'temporary' && temporaryInviteCode) {
      temporaryInviteCode.usedAtUtc = joinedAtIso
      temporaryInviteCode.usedByUserId = user.id
      temporaryInviteCode.usedByName = characterName
      temporaryInviteCode.invalidated = true
      const durationMinutes = Number(temporaryInviteCode.durationMinutes ?? 0)
      if (Number.isFinite(durationMinutes) && durationMinutes > 0) {
        profile.accessExpiresAtUtc = new Date(Date.now() + (durationMinutes * 60 * 1000)).toISOString()
      }
      markInviteAuditEntryUsed(data, code, user, joinedAtIso, 'temporary')
    } else if (codeType === 'unique') {
      markInviteAuditEntryUsed(data, code, user, joinedAtIso, 'unique')
    } else if (codeType === 'shared' && isFirstMembershipJoin) {
      const alreadyLogged = getInviteAuditLog(data).some((entry: any) => `${entry?.code ?? ''}`.trim().toUpperCase() === code && `${entry?.codeType ?? ''}`.trim().toLowerCase() === 'shared')
      if (!alreadyLogged) {
        appendInviteAuditEntry(data, {
          code,
          codeType: 'shared',
          createdByUserId: sheet.owner_id,
          createdByName: 'Sheet owner',
          createdAtUtc: sheet.created_at,
          durationMinutes: 0,
          wasUsed: true,
          usedAtUtc: joinedAtIso,
          usedByUserId: user.id,
          usedByName: characterName,
        })
      }
    }
  })

  await upsertPresenceRow(sheet.id, user.id, characterName, null, null)
  return json({ sheet_id: sheet.id })
}

async function handleDeleteSheet(req: Request, payload: any): Promise<Response> {
  const user = await requireDiscordUser(req)
  const sheetId = `${payload?.sheet_id ?? ''}`.trim()
  if (!sheetId) {
    return json({ message: 'Missing sheet_id.' }, 400)
  }

  const sheet = await getSheetOrResponse(sheetId)
  if (sheet instanceof Response) {
    return sheet
  }

  if (sheet.owner_id !== user.id) {
    return json({ message: 'Only the sheet owner can delete this sheet.' }, 403)
  }

  const { error } = await admin
    .from('sheetsync_sheets')
    .delete()
    .eq('id', sheetId)

  if (error) {
    return dbError(error)
  }

  return json({ success: true })
}

async function handleListSheetMembers(req: Request, payload: any): Promise<Response> {
  const user = await requireDiscordUser(req)
  const sheetId = `${payload?.sheet_id ?? ''}`.trim()
  if (!sheetId) {
    return json({ message: 'Missing sheet_id.' }, 400)
  }

  const sheet = await getSheetOrResponse(sheetId)
  if (sheet instanceof Response) {
    return sheet
  }

  if (sheet.owner_id !== user.id
      && !hasSheetPermissionInData(sheet, user.id, 'editPermissions')
      && !hasSheetPermissionInData(sheet, user.id, 'blockUsers')) {
    return json({ message: 'You do not have permission to view sheet members.' }, 403)
  }

  return json(await buildSheetMemberRows(sheet))
}

async function handleListSheetBlocklist(req: Request, payload: any): Promise<Response> {
  const user = await requireDiscordUser(req)
  const sheetId = `${payload?.sheet_id ?? ''}`.trim()
  if (!sheetId) {
    return json({ message: 'Missing sheet_id.' }, 400)
  }

  const sheet = await getSheetOrResponse(sheetId)
  if (sheet instanceof Response) {
    return sheet
  }

  if (sheet.owner_id !== user.id) {
    return json({ message: 'You do not have permission to view the blocklist.' }, 403)
  }

  const { data, error } = await admin
    .from('sheetsync_sheet_blocklist')
    .select('sheet_id, user_id, character_name, reason, removed_at')
    .eq('sheet_id', sheetId)
    .order('removed_at', { ascending: false })

  if (error) {
    return dbError(error)
  }

  return json(data ?? [])
}

async function handleUnblockSheetMember(req: Request, payload: any): Promise<Response> {
  const user = await requireDiscordUser(req)
  const sheetId = `${payload?.sheet_id ?? ''}`.trim()
  const userId = `${payload?.user_id ?? ''}`.trim()
  if (!sheetId || !userId) {
    return json({ message: 'Missing sheet_id or user_id.' }, 400)
  }

  const sheet = await getSheetOrResponse(sheetId)
  if (sheet instanceof Response) {
    return sheet
  }

  if (sheet.owner_id !== user.id) {
    return json({ message: 'You do not have permission to unblock members.' }, 403)
  }

  const { error } = await admin
    .from('sheetsync_sheet_blocklist')
    .delete()
    .eq('sheet_id', sheetId)
    .eq('user_id', userId)

  if (error) {
    return dbError(error)
  }

  await mutateSheetData(sheet.id, (data: any) => {
    const profile = getOrCreateProfile(data, userId)
    profile.isBlocked = false
    profile.assignedPresetName = 'Viewer'
    profile.roleColor = 0xff5c5c5c
    profile.permissions = {
      editSheet: false,
      deleteSheet: false,
      editPermissions: false,
      createTabs: false,
      seeHistory: false,
      useComments: false,
      importSheet: false,
      saveLocal: false,
      invite: false,
      blockUsers: false,
      admin: false,
    }
  })

  return json({ success: true })
}

async function handleRemoveSheetMember(req: Request, payload: any): Promise<Response> {
  const user = await requireDiscordUser(req)
  const sheetId = `${payload?.sheet_id ?? ''}`.trim()
  const userId = `${payload?.user_id ?? ''}`.trim()
  const reason = `${payload?.reason ?? ''}`.trim()
  if (!sheetId || !userId) {
    return json({ message: 'Missing sheet_id or user_id.' }, 400)
  }

  const sheet = await getSheetOrResponse(sheetId)
  if (sheet instanceof Response) {
    return sheet
  }

  if (sheet.owner_id !== user.id && !hasSheetPermissionInData(sheet, user.id, 'blockUsers')) {
    return json({ message: 'You do not have permission to remove members.' }, 403)
  }

  if (userId === sheet.owner_id) {
    return json({ message: 'The sheet owner cannot be removed.' }, 400)
  }

  const memberRows = await buildSheetMemberRows(sheet)
  const matchingMember = memberRows.find((row: any) => `${row.user_id ?? ''}` === userId)
  const characterName = `${matchingMember?.character_name ?? ''}`

  const { error } = await admin
    .from('sheetsync_sheet_members')
    .delete()
    .eq('sheet_id', sheetId)
    .eq('user_id', userId)

  if (error) {
    return dbError(error)
  }

  const { error: blockError } = await admin
    .from('sheetsync_sheet_blocklist')
    .upsert({
      sheet_id: sheetId,
      user_id: userId,
      character_name: characterName,
      reason,
      removed_at: new Date().toISOString(),
      removed_by_user_id: user.id,
    }, { onConflict: 'sheet_id,user_id' })

  if (blockError) {
    return dbError(blockError)
  }

  await admin.from('sheetsync_sheet_presence').delete().eq('sheet_id', sheetId).eq('user_id', userId)
  await admin.from('sheetsync_sheet_cell_locks').delete().eq('sheet_id', sheetId).eq('user_id', userId)

  await mutateSheetData(sheet.id, (data: any) => {
    const profile = getOrCreateProfile(data, userId)
    profile.isBlocked = true
    profile.assignedPresetName = 'Blocked'
    profile.roleColor = 0xffc84343
    profile.characterName = characterName || profile.characterName || ''
    profile.lastSeenUtc = new Date().toISOString()
    profile.permissions = {
      editSheet: false,
      deleteSheet: false,
      editPermissions: false,
      createTabs: false,
      seeHistory: false,
      useComments: false,
      importSheet: false,
      saveLocal: false,
      invite: false,
      blockUsers: false,
      admin: false,
    }
  })

  return json({ success: true })
}

async function handleSyncPresence(req: Request, payload: any): Promise<Response> {
  const user = await requireDiscordUser(req)
  const sheetId = `${payload?.sheet_id ?? ''}`.trim()
  if (!sheetId) {
    return json({ message: 'Missing sheet_id.' }, 400)
  }

  const sheet = await getSheetMetaOrResponse(sheetId)
  if (sheet instanceof Response) {
    return sheet
  }

  if (!(await getRoleForUser(sheet as any, user.id)) || await isUserBlockedFromSheet(sheetId, user.id)) {
    return json({ message: 'You do not have access to this sheet.' }, 403)
  }

  const characterName = `${payload?.character_name ?? user.global_name ?? user.username ?? 'Discord user'}`.trim()
  const activeTabName = `${payload?.active_tab_name ?? ''}`.trim()
  const editingCellKey = `${payload?.editing_cell_key ?? ''}`.trim()
  await upsertPresenceRow(sheetId, user.id, characterName, activeTabName || null, editingCellKey || null)

  return json({ success: true })
}

async function handlePostChatMessage(req: Request, payload: any): Promise<Response> {
  const user = await requireDiscordUser(req)
  const sheetId = `${payload?.sheet_id ?? ''}`.trim()
  const message = `${payload?.message ?? ''}`.trim()
  if (!sheetId || !message) {
    return json({ message: 'Missing sheet_id or message.' }, 400)
  }

  const sheet = await getSheetMetaOrResponse(sheetId)
  if (sheet instanceof Response) {
    return sheet
  }

  if (!(await getRoleForUser(sheet as any, user.id)) || await isUserBlockedFromSheet(sheetId, user.id)) {
    return json({ message: 'You do not have access to this sheet.' }, 403)
  }

  const characterName = `${payload?.character_name ?? user.global_name ?? user.username ?? 'Discord user'}`.trim()
  await purgeExpiredSheetChatRows(sheetId, getCurrentEstChatResetUtcIso())
  const { error } = await admin
    .from('sheetsync_sheet_chat_messages')
    .insert({
      sheet_id: sheetId,
      author_user_id: user.id,
      author_name: characterName,
      message,
    })

  if (error) {
    return dbError(error)
  }

  await upsertPresenceRow(sheetId, user.id, characterName, null, null)
  return json({ success: true })
}

async function handleGetSheetRuntime(req: Request, payload: any): Promise<Response> {
  const user = await requireDiscordUser(req)
  const sheetId = `${payload?.sheet_id ?? ''}`.trim()
  if (!sheetId) {
    return json({ message: 'Missing sheet_id.' }, 400)
  }

  const sheet = await getSheetMetaOrResponse(sheetId)
  if (sheet instanceof Response) {
    return sheet
  }

  if (!(await getRoleForUser(sheet as any, user.id)) || await isUserBlockedFromSheet(sheetId, user.id)) {
    return json({ message: 'You do not have access to this sheet.' }, 403)
  }

  return json(await buildSheetRuntimeState(sheet as any, user.id, Boolean(payload?.include_members ?? false)))
}

async function handleAcquireCellLock(req: Request, payload: any): Promise<Response> {
  const user = await requireDiscordUser(req)
  const sheetId = `${payload?.sheet_id ?? ''}`.trim()
  const cellKey = `${payload?.cell_key ?? ''}`.trim()
  if (!sheetId || !cellKey) {
    return json({ message: 'Missing sheet_id or cell_key.' }, 400)
  }

  const sheet = await getSheetMetaOrResponse(sheetId)
  if (sheet instanceof Response) {
    return sheet
  }

  if (!(await getRoleForUser(sheet as any, user.id)) || await isUserBlockedFromSheet(sheetId, user.id)) {
    return json({ message: 'You do not have access to this sheet.' }, 403)
  }

  const { data: existing, error: existingError } = await admin
    .from('sheetsync_sheet_cell_locks')
    .select('sheet_id, cell_key, user_id, user_name, locked_at, expires_at')
    .eq('sheet_id', sheetId)
    .eq('cell_key', cellKey)
    .gt('expires_at', new Date().toISOString())
    .maybeSingle()

  if (existingError) {
    return dbError(existingError)
  }

  if (existing && `${existing.user_id}` !== user.id) {
    return json({ message: `${existing.user_name || 'Another user'} is using this cell..` }, 409)
  }

  const characterName = `${payload?.character_name ?? user.global_name ?? user.username ?? 'Discord user'}`.trim()
  await admin
    .from('sheetsync_sheet_cell_locks')
    .delete()
    .eq('sheet_id', sheetId)
    .eq('user_id', user.id)
    .neq('cell_key', cellKey)

  const { error } = await admin
    .from('sheetsync_sheet_cell_locks')
    .upsert({
      sheet_id: sheetId,
      cell_key: cellKey,
      user_id: user.id,
      user_name: characterName,
      locked_at: new Date().toISOString(),
      expires_at: new Date(Date.now() + 15000).toISOString(),
    }, { onConflict: 'sheet_id,cell_key' })

  if (error) {
    return dbError(error)
  }

  await upsertPresenceRow(sheetId, user.id, characterName, null, cellKey)
  return json({ success: true })
}

async function handleReleaseCellLock(req: Request, payload: any): Promise<Response> {
  const user = await requireDiscordUser(req)
  const sheetId = `${payload?.sheet_id ?? ''}`.trim()
  const cellKey = `${payload?.cell_key ?? ''}`.trim()
  if (!sheetId) {
    return json({ message: 'Missing sheet_id.' }, 400)
  }

  let query = admin.from('sheetsync_sheet_cell_locks').delete().eq('sheet_id', sheetId).eq('user_id', user.id)
  if (cellKey) {
    query = query.eq('cell_key', cellKey)
  }
  const { error } = await query
  if (error) {
    return dbError(error)
  }
  return json({ success: true })
}

async function handleCreateUniqueCode(req: Request, payload: any, invalidateCurrent: boolean): Promise<Response> {
  const user = await requireDiscordUser(req)
  const sheetId = `${payload?.sheet_id ?? ''}`.trim()
  if (!sheetId) {
    return json({ message: 'Missing sheet_id.' }, 400)
  }

  const sheet = await getSheetOrResponse(sheetId)
  if (sheet instanceof Response) {
    return sheet
  }

  if (sheet.owner_id !== user.id && !hasSheetPermissionInData(sheet, user.id, 'invite')) {
    return json({ message: 'You do not have permission to manage unique codes.' }, 403)
  }

  if (invalidateCurrent) {
    const current = await getCurrentUniqueCode(sheetId)
    if (current) {
      await admin.from('sheetsync_sheet_unique_codes').update({ invalidated_at: new Date().toISOString(), invalidated_by_user_id: user.id }).eq('id', current.id)
    }
  }

  const code = generateJoinCode(12)
  const row = {
    sheet_id: sheetId,
    code,
    created_by_user_id: user.id,
  }
  const { data, error } = await admin
    .from('sheetsync_sheet_unique_codes')
    .insert(row)
    .select('id, sheet_id, code, created_at')
    .single()

  if (error) {
    return dbError(error)
  }

  await mutateSheetData(sheet.id, (dataObj: any) => {
    const dataSettings = ensureSheetData(dataObj).settings
    dataSettings.activityLog = Array.isArray(dataSettings.activityLog) ? dataSettings.activityLog : []
    dataSettings.activityLog.push({
      id: crypto.randomUUID().replace(/-/g, ''),
      userId: user.id,
      userName: `${user.global_name ?? user.username ?? 'Discord user'}`.trim(),
      action: `Generated a unique code: ${code}`,
      timestampUtc: new Date().toISOString(),
    })
    while (dataSettings.activityLog.length > 300) {
      dataSettings.activityLog.shift()
    }
    appendInviteAuditEntry(dataObj, {
      code,
      codeType: 'unique',
      createdByUserId: user.id,
      createdByName: `${user.global_name ?? user.username ?? 'Discord user'}`.trim(),
      createdAtUtc: new Date().toISOString(),
      durationMinutes: 0,
      wasUsed: false,
    })
  })

  return json(data)
}

async function getSheetOrResponse(sheetId: string): Promise<SheetRow | Response> {
  const { data, error } = await admin
    .from('sheetsync_sheets')
    .select('id, owner_id, title, code, rows_count, cols_count, default_role, version, created_at, updated_at, data')
    .eq('id', sheetId)
    .maybeSingle()

  if (error) {
    return dbError(error)
  }

  if (!data) {
    return json({ message: 'Sheet not found.' }, 404)
  }

  return data as SheetRow
}

async function getSheetMetaOrResponse(sheetId: string): Promise<SheetMetaRow | Response> {
  const { data, error } = await admin
    .from('sheetsync_sheets')
    .select('id, owner_id, title, version, updated_at')
    .eq('id', sheetId)
    .maybeSingle()

  if (error) {
    return dbError(error)
  }

  if (!data) {
    return json({ message: 'Sheet not found.' }, 404)
  }

  return data as SheetMetaRow
}

async function getRoleForUser(sheet: SheetRow, userId: string): Promise<'owner' | 'editor' | 'viewer' | null> {
  if (sheet.owner_id === userId) {
    return 'owner'
  }

  const { data, error } = await admin
    .from('sheetsync_sheet_members')
    .select('role')
    .eq('sheet_id', sheet.id)
    .eq('user_id', userId)
    .maybeSingle()

  if (error) {
    throw new Error(error.message)
  }

  if (!data?.role) {
    return null
  }

  return normalizeRole(data.role)
}

async function requireDiscordUser(req: Request): Promise<DiscordUser> {
  const serverSessionToken = `${req.headers.get('x-EZSheets-session') ?? req.headers.get('x-sheetsync-session') ?? ''}`.trim()
  if (serverSessionToken) {
    const sessionUser = await tryGetUserFromServerSession(serverSessionToken)
    if (sessionUser) {
      return sessionUser
    }
  }

  const authorization = req.headers.get('authorization') ?? ''
  const match = authorization.match(/^Bearer\s+(.+)$/i)
  if (!match) {
    throw new HttpError(401, 'Missing authorization header')
  }

  try {
    return await fetchDiscordUser(match[1])
  } catch (error) {
    if (error instanceof HttpError) {
      throw error
    }

    throw new HttpError(401, 'Discord token is invalid or expired.')
  }
}

async function exchangeDiscordToken(params: Record<string, string>): Promise<DiscordTokenResponse> {
  ensureDiscordSecrets()

  const form = new URLSearchParams()
  for (const [key, value] of Object.entries(params)) {
    form.set(key, value)
  }

  const basicAuth = btoa(`${discordClientId!}:${discordClientSecret!}`)
  const response = await fetch('https://discord.com/api/oauth2/token', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      'Authorization': `Basic ${basicAuth}`,
    },
    body: form,
  })

  const text = await response.text()
  if (!response.ok) {
    throw new HttpError(response.status, extractDiscordError(text))
  }

  return JSON.parse(text) as DiscordTokenResponse
}

async function revokeDiscordToken(token: string): Promise<void> {
  ensureDiscordSecrets()
  const form = new URLSearchParams()
  form.set('token', token)

  const basicAuth = btoa(`${discordClientId!}:${discordClientSecret!}`)
  await fetch('https://discord.com/api/oauth2/token/revoke', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      'Authorization': `Basic ${basicAuth}`,
    },
    body: form,
  })
}

async function fetchDiscordUser(accessToken: string): Promise<DiscordUser> {
  const response = await fetch('https://discord.com/api/v10/users/@me', {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
    },
  })

  const text = await response.text()
  if (!response.ok) {
    throw new HttpError(response.status, extractDiscordError(text) || 'Discord token is invalid or expired.')
  }

  return JSON.parse(text) as DiscordUser
}

async function issueServerSession(user: DiscordUser, previousToken = ''): Promise<{ token: string, expiresAtUnix: number }> {
  const token = crypto.randomUUID().replaceAll('-', '') + crypto.randomUUID().replaceAll('-', '')
  const expiresAtUnix = Math.floor(Date.now() / 1000) + (60 * 60 * 24 * 7)

  try {
    if (previousToken) {
      await revokeServerSession(previousToken)
    }

    const tokenHash = await sha256Hex(token)
    const { error } = await admin
      .from('sheetsync_discord_sessions')
      .upsert({
        token_hash: tokenHash,
        user_id: `${user.id ?? ''}`,
        user_display_name: `${user.global_name ?? user.username ?? 'Discord user'}`,
        user_email: `${user.email ?? ''}`,
        expires_at: new Date(expiresAtUnix * 1000).toISOString(),
        last_seen_at: new Date().toISOString(),
      }, { onConflict: 'token_hash' })

    if (error) {
      console.warn('Could not persist EZSheets server session.', error)
      return { token: '', expiresAtUnix: 0 }
    }

    serverSessionCache.set(tokenHash, {
      user,
      expiresAtMs: expiresAtUnix * 1000,
    })
  } catch (error) {
    console.warn('Could not issue EZSheets server session.', error)
    return { token: '', expiresAtUnix: 0 }
  }

  return { token, expiresAtUnix }
}

async function revokeServerSession(token: string): Promise<void> {
  if (!token) {
    return
  }

  try {
    const tokenHash = await sha256Hex(token)
    serverSessionCache.delete(tokenHash)
    await admin.from('sheetsync_discord_sessions').delete().eq('token_hash', tokenHash)
  } catch {
  }
}

async function tryGetUserFromServerSession(token: string): Promise<DiscordUser | null> {
  try {
    const tokenHash = await sha256Hex(token)
    const cached = serverSessionCache.get(tokenHash)
    if (cached && cached.expiresAtMs > Date.now()) {
      return cached.user
    }

    const { data, error } = await admin
      .from('sheetsync_discord_sessions')
      .select('user_id, user_display_name, user_email, expires_at')
      .eq('token_hash', tokenHash)
      .maybeSingle()

    if (error || !data) {
      serverSessionCache.delete(tokenHash)
      return null
    }

    const expiresAt = Date.parse(`${data.expires_at ?? ''}`)
    if (!Number.isFinite(expiresAt) || expiresAt <= Date.now()) {
      serverSessionCache.delete(tokenHash)
      await admin.from('sheetsync_discord_sessions').delete().eq('token_hash', tokenHash)
      return null
    }

    const user = {
      id: `${data.user_id ?? ''}`,
      global_name: `${data.user_display_name ?? ''}`,
      email: `${data.user_email ?? ''}`,
    }

    serverSessionCache.set(tokenHash, {
      user,
      expiresAtMs: expiresAt,
    })

    return user
  } catch {
    return null
  }
}

async function sha256Hex(value: string): Promise<string> {
  const bytes = new TextEncoder().encode(value)
  const digest = await crypto.subtle.digest('SHA-256', bytes)
  return Array.from(new Uint8Array(digest)).map((byte) => byte.toString(16).padStart(2, '0')).join('')
}

async function buildSession(token: DiscordTokenResponse, user: DiscordUser, previousServerSessionToken = '') {
  const serverSession = await issueServerSession(user, previousServerSessionToken)
  return {
    access_token: token.access_token,
    refresh_token: token.refresh_token ?? '',
    expires_in: token.expires_in,
    expires_at_unix: Math.floor(Date.now() / 1000) + Math.max(token.expires_in ?? 0, 0),
    token_type: token.token_type ?? 'Bearer',
    session_token: serverSession.token,
    session_expires_at_unix: serverSession.expiresAtUnix,
    user,
  }
}

function ensureDiscordSecrets() {
  if (!discordClientId || !discordClientSecret) {
    throw new Error('Missing DISCORD_CLIENT_ID or DISCORD_CLIENT_SECRET in Edge Function secrets.')
  }
}

function normalizeRole(role: unknown): 'viewer' | 'editor' {
  return `${role ?? 'viewer'}`.toLowerCase() === 'editor' ? 'editor' : 'viewer'
}

function normalizeDefaultRoleValue(role: unknown): string {
  const value = `${role ?? ''}`.trim()
  return value || 'Viewer'
}

function findPresetByName(data: any, presetName: string): any | null {
  const target = `${presetName ?? ''}`.trim().toLowerCase()
  if (!target) {
    return null
  }

  const settings = ensureSheetData(data).settings
  const presets = Array.isArray(settings.permissionPresets) ? settings.permissionPresets : []
  return presets.find((preset: any) => `${preset?.name ?? ''}`.trim().toLowerCase() === target) ?? null
}

function deriveMembershipRoleFromPermissions(permissions: any): 'viewer' | 'editor' {
  if (!permissions || typeof permissions !== 'object') {
    return 'viewer'
  }

  return permissions.editSheet
    || permissions.deleteSheet
    || permissions.editPermissions
    || permissions.createTabs
    || permissions.seeHistory
    || permissions.useComments
    || permissions.importSheet
    || permissions.saveLocal
    || permissions.invite
    || permissions.blockUsers
    || permissions.admin
    ? 'editor'
    : 'viewer'
}

function deriveMembershipRoleFromDefaultRole(defaultRole: unknown, data: any): 'viewer' | 'editor' {
  const roleName = `${defaultRole ?? ''}`.trim()
  const preset = findPresetByName(data, roleName)
  if (preset?.permissions) {
    return deriveMembershipRoleFromPermissions(preset.permissions)
  }

  return normalizeRole(roleName)
}

function applyDefaultRoleToProfile(profile: any, data: any, defaultRole: unknown): void {
  const roleName = `${defaultRole ?? ''}`.trim()
  const preset = findPresetByName(data, roleName)
  if (preset) {
    profile.assignedPresetName = (preset.name ?? roleName) || 'Viewer'
    profile.roleColor = Number(preset.color ?? profile.roleColor ?? 0xff7a7a7a)
    profile.isBlocked = false
    profile.permissions = structuredClone(preset.permissions ?? {})
    return
  }

  const normalizedRole = normalizeRole(roleName)
  profile.assignedPresetName = normalizedRole === 'editor' ? 'Editor' : 'Viewer'
  profile.roleColor = normalizedRole === 'editor' ? 0xff4f9d4f : 0xff7a7a7a
  profile.isBlocked = false
  profile.permissions = {
    editSheet: normalizedRole === 'editor',
    deleteSheet: false,
    editPermissions: false,
    createTabs: normalizedRole === 'editor',
    seeHistory: normalizedRole === 'editor',
    useComments: normalizedRole === 'editor',
    importSheet: normalizedRole === 'editor',
    saveLocal: normalizedRole === 'editor',
    invite: false,
    blockUsers: false,
    admin: false,
  }
}

function extractDiscordError(text: string): string {
  try {
    const parsed = JSON.parse(text)
    return parsed.error_description || parsed.message || parsed.error || text || 'Discord OAuth request failed.'
  } catch {
    return text || 'Discord OAuth request failed.'
  }
}

function dbError(error: { message?: string; code?: string }): Response {
  const code = `${error.code ?? ''}`
  if (code === '23505') {
    return json({ message: error.message || 'Duplicate value.' }, 409)
  }

  return json({ message: error.message || 'Database request failed.' }, 500)
}

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: jsonHeaders,
  })
}

function ensureSheetData(data: any): any {
  const normalized = data && typeof data === 'object' ? structuredClone(data) : {}
  normalized.tabs = Array.isArray(normalized.tabs) ? normalized.tabs : []
  normalized.activeTabIndex = Number.isFinite(normalized.activeTabIndex) ? normalized.activeTabIndex : 0
  normalized.settings = normalized.settings && typeof normalized.settings === 'object' ? normalized.settings : {}
  normalized.settings.presence = Array.isArray(normalized.settings.presence) ? normalized.settings.presence : []
  normalized.settings.chatMessages = Array.isArray(normalized.settings.chatMessages) ? normalized.settings.chatMessages : []
  normalized.settings.memberProfiles = Array.isArray(normalized.settings.memberProfiles) ? normalized.settings.memberProfiles : []
  normalized.settings.permissionPresets = Array.isArray(normalized.settings.permissionPresets) ? normalized.settings.permissionPresets : []
  normalized.settings.activityLog = Array.isArray(normalized.settings.activityLog) ? normalized.settings.activityLog : []
  normalized.settings.temporaryInviteCodes = Array.isArray(normalized.settings.temporaryInviteCodes) ? normalized.settings.temporaryInviteCodes : []
  normalized.settings.inviteAuditLog = Array.isArray(normalized.settings.inviteAuditLog) ? normalized.settings.inviteAuditLog : []
  return normalized
}

function getSettings(sheet: SheetRow): any {
  const data = ensureSheetData(sheet.data)
  sheet.data = data
  return data.settings
}

function getMemberProfiles(data: any): any[] {
  return ensureSheetData(data).settings.memberProfiles
}

function getPresenceEntries(data: any): any[] {
  return ensureSheetData(data).settings.presence
}

function getChatMessages(data: any): any[] {
  return ensureSheetData(data).settings.chatMessages
}

function getTemporaryInviteCodes(data: any): any[] {
  return ensureSheetData(data).settings.temporaryInviteCodes
}

function getInviteAuditLog(data: any): any[] {
  return ensureSheetData(data).settings.inviteAuditLog
}

function appendInviteAuditEntry(data: any, entry: any): void {
  const log = getInviteAuditLog(data)
  log.push({
    id: `${entry?.id ?? crypto.randomUUID().replace(/-/g, '')}`,
    code: `${entry?.code ?? ''}`.trim().toUpperCase(),
    codeType: `${entry?.codeType ?? ''}`.trim().toLowerCase(),
    createdByUserId: `${entry?.createdByUserId ?? ''}`,
    createdByName: `${entry?.createdByName ?? ''}`,
    createdAtUtc: entry?.createdAtUtc ?? new Date().toISOString(),
    durationMinutes: Number(entry?.durationMinutes ?? 0),
    wasUsed: Boolean(entry?.wasUsed ?? false),
    usedAtUtc: entry?.usedAtUtc ?? null,
    usedByUserId: `${entry?.usedByUserId ?? ''}`,
    usedByName: `${entry?.usedByName ?? ''}`,
  })
  while (log.length > 300) {
    log.shift()
  }
}

function markInviteAuditEntryUsed(data: any, code: string, user: any, usedAtIso: string, codeType?: string): void {
  const normalizedCode = `${code ?? ''}`.trim().toUpperCase()
  if (!normalizedCode) {
    return
  }

  const log = getInviteAuditLog(data)
  const targetType = `${codeType ?? ''}`.trim().toLowerCase()
  const entry = [...log].reverse().find((item: any) => `${item?.code ?? ''}`.trim().toUpperCase() === normalizedCode
    && (!targetType || `${item?.codeType ?? ''}`.trim().toLowerCase() === targetType))
  if (!entry) {
    return
  }

  entry.wasUsed = true
  entry.usedAtUtc = usedAtIso
  entry.usedByUserId = `${user?.id ?? ''}`
  entry.usedByName = `${user?.global_name ?? user?.username ?? 'Discord user'}`.trim()
}

function getOrCreateProfile(data: any, userId: string): any {
  const profiles = getMemberProfiles(data)
  let profile = profiles.find((entry: any) => `${entry?.userId ?? ''}` === userId)
  if (!profile) {
    profile = {
      userId,
      characterName: '',
      assignedPresetName: 'Viewer',
      roleColor: 0xff7a7a7a,
      isBlocked: false,
      permissions: {
        editSheet: false,
        deleteSheet: false,
        editPermissions: false,
        createTabs: false,
        seeHistory: false,
        useComments: false,
        importSheet: false,
        saveLocal: false,
        invite: false,
        blockUsers: false,
        admin: false,
      },
    }
    profiles.push(profile)
  }
  profile.permissions = profile.permissions && typeof profile.permissions === 'object' ? profile.permissions : {}
  return profile
}

function ensureOwnerProfile(data: any, ownerId: string): void {
  const profile = getOrCreateProfile(data, ownerId)
  if (!profile.joinedAtUtc) {
    profile.joinedAtUtc = new Date().toISOString()
  }
  if (!profile.assignedPresetName) {
    profile.assignedPresetName = 'Owner'
  }
}

function isUserBlocked(sheet: SheetRow, userId: string): boolean {
  if (sheet.owner_id === userId) {
    return false
  }
  const profiles = getMemberProfiles(sheet.data)
  const profile = profiles.find((entry: any) => `${entry?.userId ?? ''}` === userId)
  return Boolean(profile?.isBlocked ?? false)
}

function hasSheetPermissionInData(sheet: SheetRow, userId: string, permission: string): boolean {
  if (sheet.owner_id === userId) {
    return true
  }

  const profiles = getMemberProfiles(sheet.data)
  const profile = profiles.find((entry: any) => `${entry?.userId ?? ''}` === userId)
  if (!profile || profile.isBlocked) {
    return false
  }

  const permissions = profile.permissions ?? {}
  if (permission !== 'deleteSheet' && permissions.admin) {
    return true
  }

  return Boolean(permissions[permission])
}

async function mutateSheetData(sheetId: string, mutator: (data: any, sheet: SheetRow) => void | Promise<void>): Promise<void> {
  for (let attempt = 0; attempt < 5; attempt++) {
    const sheet = await getSheetById(sheetId)
    const data = ensureSheetData(sheet.data)
    await mutator(data, sheet)
    ensureOwnerProfile(data, sheet.owner_id)

    const { data: updated, error } = await admin
      .from('sheetsync_sheets')
      .update({ data, version: Number(sheet.version ?? 0) + 1 })
      .eq('id', sheetId)
      .eq('version', sheet.version)
      .select('id')
      .maybeSingle()

    if (error) {
      throw new Error(error.message || 'Could not save sheet data.')
    }

    if (updated) {
      return
    }
  }

  throw new Error('Could not save sheet data due to concurrent updates.')
}

function mergeSheetData(existingData: any, incomingData: any, ownerId: string): any {
  const existing = ensureSheetData(existingData)
  const incoming = ensureSheetData(incomingData)
  const merged = ensureSheetData(incoming)

  merged.settings.permissionPresets = mergeByKey(
    Array.isArray(existing.settings.permissionPresets) ? existing.settings.permissionPresets : [],
    Array.isArray(incoming.settings.permissionPresets) ? incoming.settings.permissionPresets : [],
    (item: any) => `${item?.name ?? ''}`.toLowerCase(),
  )

  merged.settings.memberProfiles = mergeMemberProfiles(
    Array.isArray(existing.settings.memberProfiles) ? existing.settings.memberProfiles : [],
    Array.isArray(incoming.settings.memberProfiles) ? incoming.settings.memberProfiles : [],
  )

  merged.settings.temporaryInviteCodes = mergeByKey(
    Array.isArray(existing.settings.temporaryInviteCodes) ? existing.settings.temporaryInviteCodes : [],
    Array.isArray(incoming.settings.temporaryInviteCodes) ? incoming.settings.temporaryInviteCodes : [],
    (item: any) => `${item?.id ?? item?.code ?? ''}`.toLowerCase(),
  )

  merged.settings.inviteAuditLog = mergeByKey(
    Array.isArray(existing.settings.inviteAuditLog) ? existing.settings.inviteAuditLog : [],
    Array.isArray(incoming.settings.inviteAuditLog) ? incoming.settings.inviteAuditLog : [],
    (item: any) => `${item?.id ?? `${item?.code ?? ''}|${item?.codeType ?? ''}`}`.toLowerCase(),
  )

  merged.settings.presence = Array.isArray(existing.settings.presence) ? existing.settings.presence : []
  merged.settings.chatMessages = Array.isArray(existing.settings.chatMessages) ? existing.settings.chatMessages : []

  ensureOwnerProfile(merged, ownerId)
  return merged
}

function mergeByKey(existing: any[], incoming: any[], getKey: (item: any) => string): any[] {
  const map = new Map<string, any>()
  for (const item of existing ?? []) {
    const key = getKey(item)
    if (key) map.set(key, structuredClone(item))
  }
  for (const item of incoming ?? []) {
    const key = getKey(item)
    if (key) map.set(key, structuredClone(item))
  }
  return Array.from(map.values())
}

function mergeMemberProfiles(existing: any[], incoming: any[]): any[] {
  const map = new Map<string, any>()
  for (const item of existing ?? []) {
    const key = `${item?.userId ?? ''}`
    if (!key) continue
    map.set(key, structuredClone(item))
  }
  for (const item of incoming ?? []) {
    const key = `${item?.userId ?? ''}`
    if (!key) continue
    const prior = map.get(key) ?? {}
    map.set(key, {
      ...prior,
      ...structuredClone(item),
      characterName: item?.characterName || prior?.characterName || '',
      joinedAtUtc: prior?.joinedAtUtc || item?.joinedAtUtc || null,
      lastSeenUtc: newerIso(prior?.lastSeenUtc, item?.lastSeenUtc),
      permissions: item?.permissions ?? prior?.permissions ?? {},
    })
  }
  return Array.from(map.values())
}

function mergePresenceEntries(existing: any[], incoming: any[]): any[] {
  const map = new Map<string, any>()
  for (const item of existing ?? []) {
    const key = `${item?.userId ?? ''}`
    if (!key) continue
    map.set(key, structuredClone(item))
  }
  for (const item of incoming ?? []) {
    const key = `${item?.userId ?? ''}`
    if (!key) continue
    const prior = map.get(key) ?? {}
    const incomingTime = Date.parse(`${item?.lastSeenUtc ?? ''}`)
    const priorTime = Date.parse(`${prior?.lastSeenUtc ?? ''}`)
    const useIncoming = Number.isFinite(incomingTime) && (!Number.isFinite(priorTime) || incomingTime >= priorTime)
    map.set(key, {
      ...prior,
      ...structuredClone(item),
      userName: item?.userName || prior?.userName || '',
      activeTabName: useIncoming ? (item?.activeTabName ?? null) : (prior?.activeTabName ?? null),
      editingCellKey: useIncoming ? (item?.editingCellKey ?? null) : (prior?.editingCellKey ?? null),
      lastSeenUtc: newerIso(prior?.lastSeenUtc, item?.lastSeenUtc),
    })
  }
  return Array.from(map.values())
}

function mergeChatMessages(existing: any[], incoming: any[]): any[] {
  const map = new Map<string, any>()
  for (const item of existing ?? []) {
    const key = `${item?.id ?? ''}` || crypto.randomUUID().replace(/-/g, '')
    map.set(key, structuredClone({ ...item, id: key }))
  }
  for (const item of incoming ?? []) {
    const key = `${item?.id ?? ''}` || crypto.randomUUID().replace(/-/g, '')
    map.set(key, structuredClone({ ...item, id: key }))
  }
  return Array.from(map.values())
    .sort((a: any, b: any) => Date.parse(`${a?.timestampUtc ?? ''}`) - Date.parse(`${b?.timestampUtc ?? ''}`))
    .slice(-200)
}

function newerIso(a: string | null | undefined, b: string | null | undefined): string | null {
  const aTime = Date.parse(`${a ?? ''}`)
  const bTime = Date.parse(`${b ?? ''}`)
  if (Number.isFinite(aTime) && (!Number.isFinite(bTime) || aTime >= bTime)) {
    return a ?? null
  }
  return b ?? null
}

async function resolveSheetByJoinCode(code: string): Promise<{ sheet: SheetRow; uniqueCodeRow: any | null; temporaryInviteCode: any | null; codeType: 'shared' | 'unique' | 'temporary' } | Response> {
  const { data: normalSheet, error: normalError } = await admin
    .from('sheetsync_sheets')
    .select('id, owner_id, title, code, rows_count, cols_count, default_role, version, created_at, updated_at, data')
    .eq('code', code)
    .maybeSingle()

  if (normalError) {
    return dbError(normalError)
  }

  if (normalSheet) {
    return { sheet: normalSheet as SheetRow, uniqueCodeRow: null, temporaryInviteCode: null, codeType: 'shared' }
  }

  const { data: uniqueCode, error: uniqueError } = await admin
    .from('sheetsync_sheet_unique_codes')
    .select('id, sheet_id, code, used_at, invalidated_at')
    .eq('code', code)
    .is('used_at', null)
    .is('invalidated_at', null)
    .maybeSingle()

  if (uniqueError) {
    return dbError(uniqueError)
  }

  if (!uniqueCode?.sheet_id) {
    return json({ message: 'No sheet exists for that code.' }, 404)
  }

  const sheet = await getSheetOrResponse(`${uniqueCode.sheet_id}`)
  if (sheet instanceof Response) {
    return sheet
  }

  const temporaryInviteCode = getTemporaryInviteCodes(sheet.data).find((entry: any) => `${entry?.code ?? ''}`.trim().toUpperCase() === code && !entry?.invalidated && !entry?.usedAtUtc) ?? null

  return { sheet, uniqueCodeRow: uniqueCode, temporaryInviteCode, codeType: temporaryInviteCode ? 'temporary' : 'unique' }
}

async function isUserBlockedFromSheet(sheetId: string, userId: string): Promise<boolean> {
  const { data, error } = await admin
    .from('sheetsync_sheet_blocklist')
    .select('user_id')
    .eq('sheet_id', sheetId)
    .eq('user_id', userId)
    .maybeSingle()

  if (error) {
    throw new Error(error.message || 'Could not check the blocklist.')
  }

  return Boolean(data?.user_id)
}

async function upsertPresenceRow(sheetId: string, userId: string, userName: string, activeTabName: string | null, editingCellKey: string | null): Promise<void> {
  const { error } = await admin
    .from('sheetsync_sheet_presence')
    .upsert({
      sheet_id: sheetId,
      user_id: userId,
      user_name: userName,
      active_tab_name: activeTabName,
      editing_cell_key: editingCellKey,
      last_seen_utc: new Date().toISOString(),
    }, { onConflict: 'sheet_id,user_id' })

  if (error) {
    throw new Error(error.message || 'Could not update presence.')
  }
}

async function buildSheetMemberRows(sheet: SheetRow): Promise<any[]> {
  const { data: memberships, error } = await admin
    .from('sheetsync_sheet_members')
    .select('sheet_id, user_id, role, created_at')
    .eq('sheet_id', sheet.id)
    .order('created_at', { ascending: true })

  if (error) {
    throw new Error(error.message || 'Could not load sheet members.')
  }

  const { data: presenceRows } = await admin
    .from('sheetsync_sheet_presence')
    .select('user_id, user_name, last_seen_utc')
    .eq('sheet_id', sheet.id)

  const presenceByUser = new Map<string, any>((presenceRows ?? []).map((row: any) => [`${row.user_id}`, row]))
  const settings = getSettings(sheet)
  const profiles = Array.isArray(settings.memberProfiles) ? settings.memberProfiles : []

  return (memberships ?? []).map((row: any) => {
    const profile = profiles.find((entry: any) => `${entry?.userId ?? ''}` === `${row.user_id}`) ?? null
    const seen = presenceByUser.get(`${row.user_id}`) ?? null
    return {
      ...row,
      character_name: `${profile?.characterName ?? seen?.user_name ?? ''}`,
      last_seen_utc: profile?.lastSeenUtc ?? seen?.last_seen_utc ?? null,
      assigned_preset_name: `${profile?.assignedPresetName ?? ''}`,
      role_color: Number(profile?.roleColor ?? 0),
      is_blocked: Boolean(profile?.isBlocked ?? false),
    }
  })
}

async function getSheetPresenceRows(sheetId: string): Promise<any[]> {
  const { data, error } = await admin
    .from('sheetsync_sheet_presence')
    .select('sheet_id, user_id, user_name, active_tab_name, editing_cell_key, last_seen_utc')
    .eq('sheet_id', sheetId)
    .gte('last_seen_utc', new Date(Date.now() - 45000).toISOString())
    .order('last_seen_utc', { ascending: false })

  if (error) {
    throw new Error(error.message || 'Could not load sheet presence.')
  }
  return data ?? []
}

function getCurrentEstChatResetUtcIso(now: Date = new Date()): string {
  const estOffsetHours = -5
  const estNow = new Date(now.getTime() + (estOffsetHours * 60 * 60 * 1000))
  const utcMidnightForEstDay = Date.UTC(estNow.getUTCFullYear(), estNow.getUTCMonth(), estNow.getUTCDate(), 5, 0, 0, 0)
  return new Date(utcMidnightForEstDay).toISOString()
}

async function purgeExpiredSheetChatRows(sheetId: string, cutoffUtcIso: string): Promise<void> {
  const { error } = await admin
    .from('sheetsync_sheet_chat_messages')
    .delete()
    .eq('sheet_id', sheetId)
    .lt('created_at', cutoffUtcIso)

  if (error) {
    throw new Error(error.message || 'Could not purge expired sheet chat messages.')
  }
}

async function getSheetChatVisibleSinceUtcIso(sheet: SheetRow, requesterUserId: string): Promise<string> {
  const dailyCutoffUtcIso = getCurrentEstChatResetUtcIso()
  if (`${sheet.owner_id ?? ''}` === requesterUserId) {
    return dailyCutoffUtcIso
  }

  const { data, error } = await admin
    .from('sheetsync_sheet_members')
    .select('created_at')
    .eq('sheet_id', sheet.id)
    .eq('user_id', requesterUserId)
    .maybeSingle()

  if (error) {
    throw new Error(error.message || 'Could not load the sheet membership timestamp.')
  }

  const joinedAtUtcIso = `${data?.created_at ?? ''}`.trim()
  if (!joinedAtUtcIso) {
    return dailyCutoffUtcIso
  }

  return new Date(joinedAtUtcIso).getTime() > new Date(dailyCutoffUtcIso).getTime()
    ? joinedAtUtcIso
    : dailyCutoffUtcIso
}

async function getSheetChatRows(sheetId: string, visibleSinceUtcIso: string): Promise<any[]> {
  await purgeExpiredSheetChatRows(sheetId, getCurrentEstChatResetUtcIso())

  const { data, error } = await admin
    .from('sheetsync_sheet_chat_messages')
    .select('id, author_user_id, author_name, message, created_at')
    .eq('sheet_id', sheetId)
    .gte('created_at', visibleSinceUtcIso)
    .order('created_at', { ascending: true })
    .limit(200)

  if (error) {
    throw new Error(error.message || 'Could not load sheet chat messages.')
  }

  return data ?? []
}

async function getSheetCellLockRows(sheetId: string): Promise<any[]> {
  const { data, error } = await admin
    .from('sheetsync_sheet_cell_locks')
    .select('sheet_id, cell_key, user_id, user_name, locked_at, expires_at')
    .eq('sheet_id', sheetId)
    .gt('expires_at', new Date().toISOString())
    .order('locked_at', { ascending: true })

  if (error) {
    throw new Error(error.message || 'Could not load cell locks.')
  }
  return data ?? []
}

async function getCurrentUniqueCode(sheetId: string): Promise<any | null> {
  const { data, error } = await admin
    .from('sheetsync_sheet_unique_codes')
    .select('id, sheet_id, code, created_at')
    .eq('sheet_id', sheetId)
    .is('used_at', null)
    .is('invalidated_at', null)
    .order('created_at', { ascending: false })
    .limit(1)
    .maybeSingle()

  if (error) {
    throw new Error(error.message || 'Could not load the current unique code.')
  }
  return data ?? null
}

async function buildSheetRuntimeState(sheet: SheetRow, requesterUserId: string, includeMembers = false): Promise<any> {
  const chatVisibleSinceUtcIso = await getSheetChatVisibleSinceUtcIso(sheet, requesterUserId)
  const [presenceRows, chatRows, memberRows, cellLocks, currentUniqueCode, requesterRole] = await Promise.all([
    getSheetPresenceRows(sheet.id),
    getSheetChatRows(sheet.id, chatVisibleSinceUtcIso),
    includeMembers ? buildSheetMemberRows(sheet) : Promise.resolve([]),
    getSheetCellLockRows(sheet.id),
    sheet.owner_id === requesterUserId
      ? getCurrentUniqueCode(sheet.id)
      : Promise.resolve(null),
    getRoleForUser(sheet, requesterUserId),
  ])

  const requesterProfile = getMemberProfiles(sheet.data).find((entry: any) => `${entry?.userId ?? ''}` === requesterUserId) ?? null
  const requesterRoleName = sheet.owner_id === requesterUserId
    ? 'Owner'
    : `${requesterProfile?.assignedPresetName ?? ''}`.trim()
  const requesterRoleColor = sheet.owner_id === requesterUserId
    ? 0xffd99a32
    : Number(requesterProfile?.roleColor ?? 0)

  return {
    sheet_id: sheet.id,
    version: Number(sheet.version ?? 0),
    updated_at: sheet.updated_at,
    presence: presenceRows.map((row: any) => ({
      userId: `${row.user_id ?? ''}`,
      userName: `${row.user_name ?? ''}`,
      activeTabName: row.active_tab_name ?? null,
      editingCellKey: row.editing_cell_key ?? null,
      lastSeenUtc: row.last_seen_utc,
    })),
    chat_messages: chatRows.map((row: any) => ({
      id: `${row.id ?? ''}`,
      authorUserId: `${row.author_user_id ?? ''}`,
      authorName: `${row.author_name ?? ''}`,
      message: `${row.message ?? ''}`,
      timestampUtc: row.created_at,
    })),
    members: memberRows,
    cell_locks: cellLocks,
    current_unique_code: currentUniqueCode,
    requester_role: `${requesterRole ?? ''}`,
    requester_role_name: requesterRoleName,
    requester_role_color: requesterRoleColor,
  }
}

function generateJoinCode(length: number): string {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'
  const values = new Uint32Array(length)
  crypto.getRandomValues(values)
  let result = ''
  for (let index = 0; index < length; index++) {
    result += alphabet[values[index] % alphabet.length]
  }
  return result
}

async function getSheetById(sheetId: string): Promise<SheetRow> {
  const sheet = await getSheetOrResponse(sheetId)
  if (sheet instanceof Response) {
    throw new Error('Sheet not found.')
  }
  return sheet
}

class HttpError extends Error {
  status: number

  constructor(status: number, message: string) {
    super(message)
    this.status = status
  }
}
