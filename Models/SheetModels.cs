using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace EZSheets.Models;

public enum SheetAccessRole
{
    Viewer,
    Editor,
    Owner,
}

public sealed class SupabaseSession
{
    [JsonPropertyName("access_token")]
    public string AccessToken { get; set; } = string.Empty;

    [JsonPropertyName("refresh_token")]
    public string RefreshToken { get; set; } = string.Empty;

    [JsonPropertyName("expires_in")]
    public int ExpiresIn { get; set; }

    [JsonPropertyName("token_type")]
    public string TokenType { get; set; } = "bearer";

    [JsonPropertyName("user")]
    public SupabaseUser? User { get; set; }
}

public sealed class SupabaseUser
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("email")]
    public string? Email { get; set; }

    [JsonPropertyName("user_metadata")]
    public Dictionary<string, JsonElement>? UserMetadata { get; set; }

    [JsonPropertyName("identities")]
    public List<SupabaseIdentity>? Identities { get; set; }

    public string GetPreferredDisplayName()
    {
        foreach (var candidate in EnumerateMetadataValues(this.UserMetadata))
        {
            if (!string.IsNullOrWhiteSpace(candidate))
            {
                return candidate;
            }
        }

        if (this.Identities is not null)
        {
            foreach (var identity in this.Identities)
            {
                foreach (var candidate in EnumerateMetadataValues(identity.IdentityData))
                {
                    if (!string.IsNullOrWhiteSpace(candidate))
                    {
                        return candidate;
                    }
                }
            }
        }

        return "Discord user";
    }

    private static IEnumerable<string> EnumerateMetadataValues(Dictionary<string, JsonElement>? source)
    {
        if (source is null)
        {
            yield break;
        }

        foreach (var key in new[] { "preferred_username", "global_name", "full_name", "name", "user_name", "username", "nick" })
        {
            if (!source.TryGetValue(key, out var value))
            {
                continue;
            }

            if (value.ValueKind == JsonValueKind.String)
            {
                var stringValue = value.GetString();
                if (!string.IsNullOrWhiteSpace(stringValue) && !stringValue.Contains("@", StringComparison.Ordinal))
                {
                    yield return stringValue;
                }
            }
        }
    }
}

public sealed class SupabaseIdentity
{
    [JsonPropertyName("identity_data")]
    public Dictionary<string, JsonElement>? IdentityData { get; set; }
}


public sealed class DiscordAppSession
{
    [JsonPropertyName("access_token")]
    public string AccessToken { get; set; } = string.Empty;

    [JsonPropertyName("refresh_token")]
    public string RefreshToken { get; set; } = string.Empty;

    [JsonPropertyName("expires_in")]
    public int ExpiresIn { get; set; }

    [JsonPropertyName("expires_at_unix")]
    public long ExpiresAtUnix { get; set; }

    [JsonPropertyName("token_type")]
    public string TokenType { get; set; } = "Bearer";

    [JsonPropertyName("session_token")]
    public string SessionToken { get; set; } = string.Empty;

    [JsonPropertyName("session_expires_at_unix")]
    public long SessionExpiresAtUnix { get; set; }

    [JsonPropertyName("user")]
    public DiscordAppUser? User { get; set; }
}

public sealed class DiscordAppUser
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("username")]
    public string? Username { get; set; }

    [JsonPropertyName("global_name")]
    public string? GlobalName { get; set; }

    [JsonPropertyName("avatar")]
    public string? Avatar { get; set; }

    [JsonPropertyName("email")]
    public string? Email { get; set; }

    public string GetPreferredDisplayName()
    {
        if (!string.IsNullOrWhiteSpace(this.GlobalName))
        {
            return this.GlobalName!;
        }

        if (!string.IsNullOrWhiteSpace(this.Username))
        {
            return this.Username!;
        }

        return "Discord user";
    }
}

public sealed class SheetCellData
{
    public string Value { get; set; } = string.Empty;

    public uint TextColor { get; set; } = 0xFFFFFFFF;

    public uint BackgroundColor { get; set; } = 0x00000000;

    public string HorizontalAlign { get; set; } = "left";

    public bool Bold { get; set; }

    public bool Italic { get; set; }

    public bool Underline { get; set; }

    public float FontScale { get; set; } = 1f;

    public bool WrapText { get; set; }

    public string VerticalAlign { get; set; } = "middle";

    public CellBorderFlags Borders { get; set; }

    public uint BorderColor { get; set; } = 0x30FFFFFF;

    public bool Locked { get; set; }

    public string? MergedInto { get; set; }

    public int RowSpan { get; set; } = 1;

    public int ColSpan { get; set; } = 1;
}

public sealed class SheetTabData
{
    public string Name { get; set; } = "Sheet";

    public uint TabColor { get; set; }

    public Dictionary<string, SheetCellData> Cells { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    public Dictionary<string, CellFeatureBundle> CellFeatures { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    public Dictionary<int, float> ColumnWidths { get; set; } = new();

    public Dictionary<int, float> RowHeights { get; set; } = new();

    public float GetColumnWidth(int column, float defaultWidth)
        => this.ColumnWidths.TryGetValue(column, out var width) && width > 0f ? width : defaultWidth;

    public float GetRowHeight(int row, float defaultHeight)
        => this.RowHeights.TryGetValue(row, out var height) && height > 0f ? height : defaultHeight;

    public SheetCellData GetOrCreateCell(string key)
    {
        if (!this.Cells.TryGetValue(key, out var cell))
        {
            cell = new SheetCellData();
            this.Cells[key] = cell;
        }

        return cell;
    }

    public bool TryGetCell(string key, out SheetCellData? cell)
    {
        var result = this.Cells.TryGetValue(key, out var found);
        cell = found;
        return result;
    }

    public CellFeatureBundle GetOrCreateFeature(string key)
    {
        if (!this.CellFeatures.TryGetValue(key, out var feature))
        {
            feature = new CellFeatureBundle { CellKey = key };
            this.CellFeatures[key] = feature;
        }

        return feature;
    }

    public bool TryGetFeature(string key, out CellFeatureBundle? feature)
    {
        var result = this.CellFeatures.TryGetValue(key, out var found);
        feature = found;
        return result;
    }
}

public sealed class SheetDocument
{
    public List<SheetTabData> Tabs { get; set; } = new();

    public int ActiveTabIndex { get; set; }

    public ExpandedSheetSettings Settings { get; set; } = new();

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public Dictionary<string, SheetCellData>? Cells { get; set; }

    public void Normalize(int rowsCount, int colsCount)
    {
        if (this.Tabs.Count == 0)
        {
            var firstTab = new SheetTabData
            {
                Name = "Sheet 1",
                Cells = this.Cells is not null
                    ? new Dictionary<string, SheetCellData>(this.Cells, StringComparer.OrdinalIgnoreCase)
                    : new Dictionary<string, SheetCellData>(StringComparer.OrdinalIgnoreCase),
            };

            this.Tabs.Add(firstTab);
        }

        this.Cells = null;

        this.Settings ??= new ExpandedSheetSettings();

        for (var index = 0; index < this.Tabs.Count; index++)
        {
            var tab = this.Tabs[index];
            if (string.IsNullOrWhiteSpace(tab.Name))
            {
                tab.Name = $"Sheet {index + 1}";
            }

            tab.Cells ??= new Dictionary<string, SheetCellData>(StringComparer.OrdinalIgnoreCase);
            tab.CellFeatures ??= new Dictionary<string, CellFeatureBundle>(StringComparer.OrdinalIgnoreCase);
        }

        if (this.ActiveTabIndex < 0 || this.ActiveTabIndex >= this.Tabs.Count)
        {
            this.ActiveTabIndex = 0;
        }

        _ = rowsCount;
        _ = colsCount;
    }

    public SheetTabData GetActiveTab(int rowsCount, int colsCount)
    {
        this.Normalize(rowsCount, colsCount);
        return this.Tabs[this.ActiveTabIndex];
    }

    public SheetTabData AddNewTab(string? name = null)
    {
        var index = this.Tabs.Count + 1;
        var tab = new SheetTabData { Name = string.IsNullOrWhiteSpace(name) ? $"Sheet {index}" : name.Trim() };
        this.Tabs.Add(tab);
        this.ActiveTabIndex = this.Tabs.Count - 1;
        return tab;
    }

    public static SheetDocument CreateDefault()
    {
        return new SheetDocument
        {
            Tabs = new List<SheetTabData>
            {
                new() { Name = "Sheet 1" },
                new() { Name = "Sheet 2" },
            },
            ActiveTabIndex = 0,
        };
    }
}

public class SheetSummary
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("owner_id")]
    public string OwnerId { get; set; } = string.Empty;

    [JsonPropertyName("title")]
    public string Title { get; set; } = string.Empty;

    [JsonPropertyName("code")]
    public string Code { get; set; } = string.Empty;

    [JsonPropertyName("rows_count")]
    public int RowsCount { get; set; }

    [JsonPropertyName("cols_count")]
    public int ColsCount { get; set; }

    [JsonPropertyName("default_role")]
    public string DefaultRole { get; set; } = "viewer";

    [JsonPropertyName("version")]
    public long Version { get; set; }

    [JsonPropertyName("created_at")]
    public DateTimeOffset CreatedAt { get; set; }

    [JsonPropertyName("updated_at")]
    public DateTimeOffset UpdatedAt { get; set; }

    [JsonPropertyName("user_role")]
    public string UserRole { get; set; } = string.Empty;

    [JsonPropertyName("user_role_color")]
    public uint UserRoleColor { get; set; }
}

public sealed class RemoteSheet : SheetSummary
{
    [JsonPropertyName("data")]
    public SheetDocument Data { get; set; } = new();
}

public sealed class SheetMemberRow
{
    [JsonPropertyName("sheet_id")]
    public string SheetId { get; set; } = string.Empty;

    [JsonPropertyName("user_id")]
    public string UserId { get; set; } = string.Empty;

    [JsonPropertyName("role")]
    public string Role { get; set; } = "viewer";

    [JsonPropertyName("created_at")]
    public DateTimeOffset CreatedAt { get; set; }

    [JsonPropertyName("character_name")]
    public string CharacterName { get; set; } = string.Empty;

    [JsonPropertyName("last_seen_utc")]
    public DateTimeOffset? LastSeenUtc { get; set; }

    [JsonPropertyName("assigned_preset_name")]
    public string AssignedPresetName { get; set; } = string.Empty;

    [JsonPropertyName("role_color")]
    public uint RoleColor { get; set; }

    [JsonPropertyName("is_blocked")]
    public bool IsBlocked { get; set; }
}



public sealed class SheetCellEditLock
{
    [JsonPropertyName("sheet_id")]
    public string SheetId { get; set; } = string.Empty;

    [JsonPropertyName("cell_key")]
    public string CellKey { get; set; } = string.Empty;

    [JsonPropertyName("user_id")]
    public string UserId { get; set; } = string.Empty;

    [JsonPropertyName("user_name")]
    public string UserName { get; set; } = string.Empty;

    [JsonPropertyName("locked_at")]
    public DateTimeOffset LockedAt { get; set; }

    [JsonPropertyName("expires_at")]
    public DateTimeOffset ExpiresAt { get; set; }
}

public sealed class SheetUniqueCodeInfo
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("sheet_id")]
    public string SheetId { get; set; } = string.Empty;

    [JsonPropertyName("code")]
    public string Code { get; set; } = string.Empty;

    [JsonPropertyName("created_at")]
    public DateTimeOffset CreatedAt { get; set; }
}

public sealed class SheetBlockedMemberRow
{
    [JsonPropertyName("sheet_id")]
    public string SheetId { get; set; } = string.Empty;

    [JsonPropertyName("user_id")]
    public string UserId { get; set; } = string.Empty;

    [JsonPropertyName("character_name")]
    public string CharacterName { get; set; } = string.Empty;

    [JsonPropertyName("reason")]
    public string Reason { get; set; } = string.Empty;

    [JsonPropertyName("removed_at")]
    public DateTimeOffset RemovedAt { get; set; }
}

public sealed class SheetRuntimeState
{
    [JsonPropertyName("sheet_id")]
    public string SheetId { get; set; } = string.Empty;

    [JsonPropertyName("version")]
    public long Version { get; set; }

    [JsonPropertyName("updated_at")]
    public DateTimeOffset UpdatedAt { get; set; }

    [JsonPropertyName("presence")]
    public List<SheetPresenceInfo> Presence { get; set; } = new();

    [JsonPropertyName("chat_messages")]
    public List<SheetChatMessage> ChatMessages { get; set; } = new();

    [JsonPropertyName("members")]
    public List<SheetMemberRow> Members { get; set; } = new();

    [JsonPropertyName("cell_locks")]
    public List<SheetCellEditLock> CellLocks { get; set; } = new();

    [JsonPropertyName("current_unique_code")]
    public SheetUniqueCodeInfo? CurrentUniqueCode { get; set; }

    [JsonPropertyName("requester_role")]
    public string RequesterRole { get; set; } = string.Empty;

    [JsonPropertyName("requester_role_name")]
    public string RequesterRoleName { get; set; } = string.Empty;

    [JsonPropertyName("requester_role_color")]
    public uint RequesterRoleColor { get; set; }
}

public sealed class LocalSheetCache
{
    public string SheetId { get; set; } = string.Empty;

    public string Title { get; set; } = string.Empty;

    public string Code { get; set; } = string.Empty;

    public int RowsCount { get; set; }

    public int ColsCount { get; set; }

    public string DefaultRole { get; set; } = "viewer";

    public long Version { get; set; }

    public SheetDocument Data { get; set; } = new();

    public DateTime LastLocalSaveUtc { get; set; } = DateTime.UtcNow;
}

public sealed class ApiErrorResponse
{
    [JsonPropertyName("message")]
    public string? Message { get; set; }

    [JsonPropertyName("msg")]
    public string? Msg { get; set; }

    [JsonPropertyName("error_description")]
    public string? ErrorDescription { get; set; }

    [JsonPropertyName("error")]
    public string? Error { get; set; }

    public string ToFriendlyMessage()
    {
        return this.ErrorDescription
               ?? this.Message
               ?? this.Msg
               ?? this.Error
               ?? "Supabase returned an unknown error.";
    }
}
