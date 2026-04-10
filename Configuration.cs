using Dalamud.Configuration;
using EZSheets.Models;

namespace EZSheets;

[Serializable]
public sealed class Configuration : IPluginConfiguration
{
    public const string EmbeddedSupabaseUrl = "https://ngcuigglmccqtwfcwwar.supabase.co";
    public const string EmbeddedSupabasePublishableKey = "sb_publishable_r0U-qMwhRLgHPit7RcFnRg_qPn0VWNQ";
    public const string EmbeddedDiscordClientId = "1490748744625557695";
    public const int DiscordCallbackPort = 38473;
    public const string DiscordCallbackUrl = "http://127.0.0.1:38473/callback/";
    public const string EZSheetsFunctionName = "EZSheets-api";
    public const string LegacySupabaseFunctionName = "sheetsync-api";

    public int Version { get; set; } = 5;

    public string SupabaseUrl { get; set; } = EmbeddedSupabaseUrl;

    public string SupabasePublicApiKey { get; set; } = EmbeddedSupabasePublishableKey;

    public string DiscordClientId { get; set; } = EmbeddedDiscordClientId;

    public string AccessToken { get; set; } = string.Empty;

    public string RefreshToken { get; set; } = string.Empty;

    public long AccessTokenExpiresAtUnix { get; set; }

    public string ServerSessionToken { get; set; } = string.Empty;

    public long ServerSessionExpiresAtUnix { get; set; }

    public string RememberedRefreshToken { get; set; } = string.Empty;

    public string UserId { get; set; } = string.Empty;

    public string UserEmail { get; set; } = string.Empty;

    public string UserDisplayName { get; set; } = string.Empty;

    public string LastCharacterFullName { get; set; } = string.Empty;

    public bool SidebarCollapsed { get; set; }

    public float MainWindowPosX { get; set; }

    public float MainWindowPosY { get; set; }

    public float MainWindowWidth { get; set; }

    public float MainWindowHeight { get; set; }

    public bool HasSavedMainWindowLayout { get; set; }

    public string LastOpenedSheetId { get; set; } = string.Empty;

    public string LastLocalExportPath { get; set; } = string.Empty;

    public Dictionary<string, LocalSheetCache> LocalSheets { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    public Dictionary<string, uint> SheetListColors { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    public uint LastBorderColor { get; set; } = 0xCCFFFFFF;

    public Dictionary<string, CharacterSessionSlot> CharacterSessions { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, long> RecentSheetAccessTicks { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    public float ChatWindowPosX { get; set; }

    public float ChatWindowPosY { get; set; }

    public float ChatWindowWidth { get; set; }

    public float ChatWindowHeight { get; set; }

    public bool HasSavedChatWindowLayout { get; set; }

    public float SheetListWindowPosX { get; set; }

    public float SheetListWindowPosY { get; set; }

    public float SheetListWindowWidth { get; set; }

    public float SheetListWindowHeight { get; set; }

    public bool HasSavedSheetListWindowLayout { get; set; }

    public float PresenceWindowPosX { get; set; }

    public float PresenceWindowPosY { get; set; }

    public float PresenceWindowWidth { get; set; }

    public float PresenceWindowHeight { get; set; }

    public bool HasSavedPresenceWindowLayout { get; set; }

    public Dictionary<string, bool> SheetFavoriteStates { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    public Dictionary<string, string> SheetChatReadAnchors { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    public Dictionary<string, string> SheetChatMentionToastAnchors { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    public void Save()
    {
        Plugin.PluginInterface.SavePluginConfig(this);
    }

    public void ClearSession(bool keepRememberedLogin = true)
    {
        this.AccessToken = string.Empty;
        this.RefreshToken = string.Empty;
        this.AccessTokenExpiresAtUnix = 0;
        this.ServerSessionToken = string.Empty;
        this.ServerSessionExpiresAtUnix = 0;
        this.UserId = string.Empty;
        this.UserEmail = string.Empty;
        this.UserDisplayName = string.Empty;
        if (!keepRememberedLogin)
        {
            this.RememberedRefreshToken = string.Empty;
            this.LastCharacterFullName = string.Empty;
        }
    }

    public bool TryLoadCharacterSession(string? characterFullName)
    {
        this.CharacterSessions ??= new Dictionary<string, CharacterSessionSlot>(StringComparer.OrdinalIgnoreCase);
        this.RecentSheetAccessTicks ??= new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        this.SheetFavoriteStates ??= new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        this.SheetChatReadAnchors ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        this.SheetChatMentionToastAnchors ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        var key = NormalizeCharacterSessionKey(characterFullName);
        if (string.IsNullOrWhiteSpace(key) || !this.CharacterSessions.TryGetValue(key, out var slot))
        {
            this.ClearSession();
            return false;
        }

        this.AccessToken = slot.AccessToken ?? string.Empty;
        this.RefreshToken = slot.RefreshToken ?? string.Empty;
        this.AccessTokenExpiresAtUnix = slot.AccessTokenExpiresAtUnix;
        this.ServerSessionToken = slot.ServerSessionToken ?? string.Empty;
        this.ServerSessionExpiresAtUnix = slot.ServerSessionExpiresAtUnix;
        this.RememberedRefreshToken = slot.RememberedRefreshToken ?? string.Empty;
        this.UserId = slot.UserId ?? string.Empty;
        this.UserEmail = slot.UserEmail ?? string.Empty;
        this.UserDisplayName = slot.UserDisplayName ?? string.Empty;
        this.LastCharacterFullName = string.IsNullOrWhiteSpace(slot.LastCharacterFullName)
            ? (characterFullName?.Trim() ?? string.Empty)
            : slot.LastCharacterFullName.Trim();
        return true;
    }

    public bool TryReloadCharacterSessionFromPersisted(string? characterFullName)
    {
        this.CharacterSessions ??= new Dictionary<string, CharacterSessionSlot>(StringComparer.OrdinalIgnoreCase);
        this.RecentSheetAccessTicks ??= new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        this.SheetFavoriteStates ??= new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        this.SheetChatReadAnchors ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        this.SheetChatMentionToastAnchors ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        var key = NormalizeCharacterSessionKey(characterFullName);
        if (string.IsNullOrWhiteSpace(key))
        {
            return false;
        }

        try
        {
            if (Plugin.PluginInterface.GetPluginConfig() is Configuration persisted
                && !ReferenceEquals(persisted, this)
                && persisted.CharacterSessions is not null
                && persisted.CharacterSessions.TryGetValue(key, out var slot)
                && slot is not null)
            {
                this.CharacterSessions[key] = slot;
            }
        }
        catch
        {
            // Best-effort reload only.
        }

        return this.TryLoadCharacterSession(characterFullName);
    }

    public void SaveCurrentSessionToCharacter(string? characterFullName)
    {
        this.CharacterSessions ??= new Dictionary<string, CharacterSessionSlot>(StringComparer.OrdinalIgnoreCase);
        this.RecentSheetAccessTicks ??= new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        this.SheetFavoriteStates ??= new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        this.SheetChatReadAnchors ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        this.SheetChatMentionToastAnchors ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        this.MergePersistedCharacterSessions();

        var key = NormalizeCharacterSessionKey(characterFullName);
        if (string.IsNullOrWhiteSpace(key))
        {
            return;
        }

        var fullName = string.IsNullOrWhiteSpace(characterFullName)
            ? this.LastCharacterFullName?.Trim() ?? string.Empty
            : characterFullName.Trim();

        var slot = new CharacterSessionSlot
        {
            CharacterFullName = fullName,
            LastCharacterFullName = fullName,
            AccessToken = this.AccessToken ?? string.Empty,
            RefreshToken = this.RefreshToken ?? string.Empty,
            AccessTokenExpiresAtUnix = this.AccessTokenExpiresAtUnix,
            ServerSessionToken = this.ServerSessionToken ?? string.Empty,
            ServerSessionExpiresAtUnix = this.ServerSessionExpiresAtUnix,
            RememberedRefreshToken = string.IsNullOrWhiteSpace(this.RememberedRefreshToken)
                ? (this.RefreshToken ?? string.Empty)
                : this.RememberedRefreshToken,
            UserId = this.UserId ?? string.Empty,
            UserEmail = this.UserEmail ?? string.Empty,
            UserDisplayName = this.UserDisplayName ?? string.Empty,
        };

        this.CharacterSessions[key] = slot;
        if (!string.IsNullOrWhiteSpace(fullName))
        {
            this.LastCharacterFullName = fullName;
        }
    }

    public void ClearCharacterSession(string? characterFullName, bool keepRememberedLogin = true)
    {
        this.CharacterSessions ??= new Dictionary<string, CharacterSessionSlot>(StringComparer.OrdinalIgnoreCase);
        this.RecentSheetAccessTicks ??= new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        this.SheetFavoriteStates ??= new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        this.SheetChatReadAnchors ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        this.SheetChatMentionToastAnchors ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        this.MergePersistedCharacterSessions();

        var key = NormalizeCharacterSessionKey(characterFullName);
        if (!string.IsNullOrWhiteSpace(key) && this.CharacterSessions.TryGetValue(key, out var slot))
        {
            slot.AccessToken = string.Empty;
            slot.RefreshToken = string.Empty;
            slot.AccessTokenExpiresAtUnix = 0;
            slot.ServerSessionToken = string.Empty;
            slot.ServerSessionExpiresAtUnix = 0;
            slot.UserId = string.Empty;
            slot.UserEmail = string.Empty;
            slot.UserDisplayName = string.Empty;
            if (!keepRememberedLogin)
            {
                slot.RememberedRefreshToken = string.Empty;
            }

            this.CharacterSessions[key] = slot;
        }

        this.ClearSession(keepRememberedLogin);
    }

    public static string NormalizeCharacterSessionKey(string? characterFullName)
    {
        return string.IsNullOrWhiteSpace(characterFullName)
            ? string.Empty
            : characterFullName.Trim().ToLowerInvariant();
    }

    private void MergePersistedCharacterSessions()
    {
        try
        {
            if (Plugin.PluginInterface.GetPluginConfig() is not Configuration persisted
                || ReferenceEquals(persisted, this)
                || persisted.CharacterSessions is null)
            {
                return;
            }

            foreach (var entry in persisted.CharacterSessions)
            {
                if (!this.CharacterSessions.TryGetValue(entry.Key, out var existing))
                {
                    this.CharacterSessions[entry.Key] = entry.Value;
                    continue;
                }

                var persistedSlot = entry.Value;
                var persistedLooksNewer = persistedSlot.ServerSessionExpiresAtUnix >= existing.ServerSessionExpiresAtUnix
                    || persistedSlot.AccessTokenExpiresAtUnix >= existing.AccessTokenExpiresAtUnix
                    || (!string.IsNullOrWhiteSpace(persistedSlot.ServerSessionToken) && string.IsNullOrWhiteSpace(existing.ServerSessionToken))
                    || (!string.IsNullOrWhiteSpace(persistedSlot.AccessToken) && string.IsNullOrWhiteSpace(existing.AccessToken))
                    || (!string.IsNullOrWhiteSpace(persistedSlot.RefreshToken) && string.IsNullOrWhiteSpace(existing.RefreshToken))
                    || (!string.IsNullOrWhiteSpace(persistedSlot.RememberedRefreshToken) && string.IsNullOrWhiteSpace(existing.RememberedRefreshToken));

                if (persistedLooksNewer)
                {
                    this.CharacterSessions[entry.Key] = persistedSlot;
                }
            }
        }
        catch
        {
            // Best-effort merge only.
        }
    }

    public void EnsureEmbeddedDefaults()
    {
        this.SupabaseUrl = NormalizeSupabaseUrl(this.SupabaseUrl);

        if (string.IsNullOrWhiteSpace(this.SupabasePublicApiKey) || !this.SupabasePublicApiKey.StartsWith("sb_publishable_", StringComparison.OrdinalIgnoreCase))
        {
            this.SupabasePublicApiKey = EmbeddedSupabasePublishableKey;
        }

        if (string.IsNullOrWhiteSpace(this.DiscordClientId))
        {
            this.DiscordClientId = EmbeddedDiscordClientId;
        }

        this.LocalSheets ??= new Dictionary<string, LocalSheetCache>(StringComparer.OrdinalIgnoreCase);
        this.SheetListColors ??= new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);
        this.CharacterSessions ??= new Dictionary<string, CharacterSessionSlot>(StringComparer.OrdinalIgnoreCase);
        this.RecentSheetAccessTicks ??= new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        this.SheetFavoriteStates ??= new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        this.SheetChatReadAnchors ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        this.SheetChatMentionToastAnchors ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        if (!string.IsNullOrWhiteSpace(this.LastCharacterFullName)
            && (!string.IsNullOrWhiteSpace(this.RefreshToken)
                || !string.IsNullOrWhiteSpace(this.RememberedRefreshToken)
                || !string.IsNullOrWhiteSpace(this.AccessToken))
            && !this.CharacterSessions.ContainsKey(NormalizeCharacterSessionKey(this.LastCharacterFullName)))
        {
            this.SaveCurrentSessionToCharacter(this.LastCharacterFullName);
        }

        if (this.LastBorderColor == 0)
        {
            this.LastBorderColor = 0xCCFFFFFF;
        }
    }

    public static string NormalizeSupabaseUrl(string? rawUrl)
    {
        var candidate = rawUrl?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(candidate))
        {
            return EmbeddedSupabaseUrl;
        }

        if (candidate.StartsWith("ttps://", StringComparison.OrdinalIgnoreCase))
        {
            candidate = "h" + candidate;
        }
        else if (candidate.StartsWith("tps://", StringComparison.OrdinalIgnoreCase))
        {
            candidate = "ht" + candidate;
        }
        else if (!candidate.Contains("://", StringComparison.Ordinal))
        {
            candidate = $"https://{candidate}";
        }

        if (Uri.TryCreate(candidate, UriKind.Absolute, out var parsed)
            && (parsed.Scheme == Uri.UriSchemeHttps || parsed.Scheme == Uri.UriSchemeHttp))
        {
            return parsed.ToString().TrimEnd('/');
        }

        return EmbeddedSupabaseUrl;
    }
}


[Serializable]
public sealed class CharacterSessionSlot
{
    public string CharacterFullName { get; set; } = string.Empty;

    public string LastCharacterFullName { get; set; } = string.Empty;

    public string AccessToken { get; set; } = string.Empty;

    public string RefreshToken { get; set; } = string.Empty;

    public long AccessTokenExpiresAtUnix { get; set; }

    public string ServerSessionToken { get; set; } = string.Empty;

    public long ServerSessionExpiresAtUnix { get; set; }

    public string RememberedRefreshToken { get; set; } = string.Empty;

    public string UserId { get; set; } = string.Empty;

    public string UserEmail { get; set; } = string.Empty;

    public string UserDisplayName { get; set; } = string.Empty;
}
