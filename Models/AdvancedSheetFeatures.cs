using System.Text.Json.Serialization;

namespace EZSheets.Models;

public enum SheetTemplateKind
{
    Blank,
    Schedule,
    Staff,
    Bartender,
    Waitlist,
    Blackjack,
    VenueAddressBook,
    Finance,
}

public enum SheetViewMode
{
    Standard,
    Presentation,
    ReadOnly,
}

public enum CellVerticalAlign
{
    Top,
    Middle,
    Bottom,
}

[Flags]
public enum CellBorderFlags
{
    None = 0,
    Left = 1 << 0,
    Top = 1 << 1,
    Right = 1 << 2,
    Bottom = 1 << 3,
    All = Left | Top | Right | Bottom,
}

public enum ConditionalOperator
{
    Equal,
    NotEqual,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
    Contains,
    StartsWith,
    EndsWith,
}

public sealed class CellComment
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    public string AuthorId { get; set; } = string.Empty;

    public string AuthorName { get; set; } = string.Empty;

    public string Message { get; set; } = string.Empty;

    public DateTimeOffset CreatedAtUtc { get; set; } = DateTimeOffset.UtcNow;
}

public sealed class DropdownValidation
{
    public List<string> Options { get; set; } = new();

    public bool AllowCustomText { get; set; } = true;
}

public sealed class ProtectedRange
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    public int StartRow { get; set; }

    public int StartColumn { get; set; }

    public int EndRow { get; set; }

    public int EndColumn { get; set; }

    public List<string> EditorUserIds { get; set; } = new();

    public bool AllowOwnerOnly { get; set; }
}

public sealed class ConditionalFormattingRule
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    public ConditionalOperator Operator { get; set; }

    public string TargetValue { get; set; } = string.Empty;

    public uint? TextColor { get; set; }

    public uint? BackgroundColor { get; set; }

    public bool? Bold { get; set; }

    public bool? Italic { get; set; }

    public bool? Underline { get; set; }
}

public sealed class ActivityEntry
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    public string UserId { get; set; } = string.Empty;

    public string UserName { get; set; } = string.Empty;

    public string Action { get; set; } = string.Empty;

    public string? CellKey { get; set; }

    public string? TabName { get; set; }

    public string? OldValue { get; set; }

    public string? NewValue { get; set; }

    public DateTimeOffset TimestampUtc { get; set; } = DateTimeOffset.UtcNow;
}

public sealed class SheetVersionSnapshot
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    public string Label { get; set; } = "Snapshot";

    public string CreatedByUserId { get; set; } = string.Empty;

    public string CreatedByName { get; set; } = string.Empty;

    public DateTimeOffset CreatedAtUtc { get; set; } = DateTimeOffset.UtcNow;

    public string DocumentJson { get; set; } = string.Empty;
}

public sealed class SheetPresenceInfo
{
    public string UserId { get; set; } = string.Empty;

    public string UserName { get; set; } = string.Empty;

    public string? ActiveTabName { get; set; }

    public string? EditingCellKey { get; set; }

    public DateTimeOffset LastSeenUtc { get; set; } = DateTimeOffset.UtcNow;
}

public sealed class SheetChatMessage
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    public string AuthorUserId { get; set; } = string.Empty;

    public string AuthorName { get; set; } = string.Empty;

    public string Message { get; set; } = string.Empty;

    public DateTimeOffset TimestampUtc { get; set; } = DateTimeOffset.UtcNow;
}


public enum SheetPermissionType
{
    EditSheet,
    DeleteSheet,
    EditPermissions,
    CreateTabs,
    SeeHistory,
    UseComments,
    ImportSheet,
    SaveLocal,
    Invite,
    BlockUsers,
    Admin,
}

public sealed class SheetPermissionSet
{
    public bool EditSheet { get; set; }

    public bool DeleteSheet { get; set; }

    public bool EditPermissions { get; set; }

    public bool CreateTabs { get; set; }

    public bool SeeHistory { get; set; }

    public bool UseComments { get; set; }

    public bool ImportSheet { get; set; }

    public bool SaveLocal { get; set; }

    public bool Invite { get; set; }

    public bool BlockUsers { get; set; }

    public bool Admin { get; set; }

    public bool Has(SheetPermissionType permission)
    {
        if (this.Admin && permission != SheetPermissionType.DeleteSheet)
        {
            return true;
        }

        return permission switch
        {
            SheetPermissionType.EditSheet => this.EditSheet,
            SheetPermissionType.DeleteSheet => this.DeleteSheet,
            SheetPermissionType.EditPermissions => this.EditPermissions,
            SheetPermissionType.CreateTabs => this.CreateTabs,
            SheetPermissionType.SeeHistory => this.SeeHistory,
            SheetPermissionType.UseComments => this.UseComments,
            SheetPermissionType.ImportSheet => this.ImportSheet,
            SheetPermissionType.SaveLocal => this.SaveLocal,
            SheetPermissionType.Invite => this.Invite,
            SheetPermissionType.BlockUsers => this.BlockUsers,
            SheetPermissionType.Admin => this.Admin,
            _ => false,
        };
    }

    public void CopyFrom(SheetPermissionSet? other)
    {
        if (other is null)
        {
            this.Reset();
            return;
        }

        this.EditSheet = other.EditSheet;
        this.DeleteSheet = other.DeleteSheet;
        this.EditPermissions = other.EditPermissions;
        this.CreateTabs = other.CreateTabs;
        this.SeeHistory = other.SeeHistory;
        this.UseComments = other.UseComments;
        this.ImportSheet = other.ImportSheet;
        this.SaveLocal = other.SaveLocal;
        this.Invite = other.Invite;
        this.BlockUsers = other.BlockUsers;
        this.Admin = other.Admin;
    }

    public void Reset()
    {
        this.EditSheet = false;
        this.DeleteSheet = false;
        this.EditPermissions = false;
        this.CreateTabs = false;
        this.SeeHistory = false;
        this.UseComments = false;
        this.ImportSheet = false;
        this.SaveLocal = false;
        this.Invite = false;
        this.BlockUsers = false;
        this.Admin = false;
    }

    public int CountAllowed()
        => new[]
        {
            this.EditSheet,
            this.DeleteSheet,
            this.EditPermissions,
            this.CreateTabs,
            this.SeeHistory,
            this.UseComments,
            this.ImportSheet,
            this.SaveLocal,
            this.Invite,
            this.BlockUsers,
            this.Admin,
        }.Count(x => x);

    public IEnumerable<string> EnumerateAllowedLabels()
    {
        if (this.Admin)
        {
            yield return "Admin";
        }

        if (this.EditSheet)
        {
            yield return "Edit Sheet";
        }

        if (this.DeleteSheet)
        {
            yield return "Delete Sheet";
        }

        if (this.EditPermissions)
        {
            yield return "Edit Role";
        }

        if (this.CreateTabs)
        {
            yield return "Create Tabs";
        }

        if (this.SeeHistory)
        {
            yield return "See History";
        }

        if (this.UseComments)
        {
            yield return "Use Comments";
        }

        if (this.ImportSheet)
        {
            yield return "Import Sheet";
        }

        if (this.SaveLocal)
        {
            yield return "Save Local";
        }

        if (this.Invite)
        {
            yield return "Invite";
        }

        if (this.BlockUsers)
        {
            yield return "Block Users";
        }
    }
}

public sealed class SheetPermissionPreset
{
    public string Name { get; set; } = "Viewer";

    public uint Color { get; set; } = 0xFF9A9A9A;

    public SheetPermissionSet Permissions { get; set; } = new();
}

public sealed class SheetTabPreset
{
    public string Name { get; set; } = "Preset";

    public SheetTabData Tab { get; set; } = new();
}

public sealed class SheetMemberProfile
{
    public string UserId { get; set; } = string.Empty;

    public string CharacterName { get; set; } = string.Empty;

    public DateTimeOffset? JoinedAtUtc { get; set; }

    public DateTimeOffset? LastSeenUtc { get; set; }

    public DateTimeOffset? AccessExpiresAtUtc { get; set; }

    public string AssignedPresetName { get; set; } = "Viewer";

    public uint RoleColor { get; set; } = 0xFF9A9A9A;

    public bool IsBlocked { get; set; }

    public SheetPermissionSet Permissions { get; set; } = new();
}


public sealed class TemporaryInviteCode
{
    public string Id { get; set; } = string.Empty;

    public string Code { get; set; } = string.Empty;

    public string CreatedByUserId { get; set; } = string.Empty;

    public string CreatedByName { get; set; } = string.Empty;

    public DateTimeOffset CreatedAtUtc { get; set; }

    public int DurationMinutes { get; set; }

    public DateTimeOffset? UsedAtUtc { get; set; }

    public string UsedByUserId { get; set; } = string.Empty;

    public string UsedByName { get; set; } = string.Empty;

    public bool Invalidated { get; set; }
}

public sealed class InviteAuditEntry
{
    public string Id { get; set; } = string.Empty;

    public string Code { get; set; } = string.Empty;

    public string CodeType { get; set; } = string.Empty;

    public string CreatedByUserId { get; set; } = string.Empty;

    public string CreatedByName { get; set; } = string.Empty;

    public DateTimeOffset CreatedAtUtc { get; set; }

    public int DurationMinutes { get; set; }

    public bool WasUsed { get; set; }

    public DateTimeOffset? UsedAtUtc { get; set; }

    public string UsedByUserId { get; set; } = string.Empty;

    public string UsedByName { get; set; } = string.Empty;
}

public sealed class ExpandedSheetSettings
{
    public SheetTemplateKind TemplateKind { get; set; } = SheetTemplateKind.Blank;

    public SheetViewMode ViewMode { get; set; } = SheetViewMode.Standard;

    public string Category { get; set; } = string.Empty;

    public bool Favorite { get; set; }

    public bool AutoSaveLocal { get; set; }

    public bool AutoSaveCloud { get; set; } = true;

    public int AutoSaveSeconds { get; set; } = 30;

    public int FrozenRows { get; set; }

    public int FrozenColumns { get; set; }

    public string ThemeName { get; set; } = "Default";

    public List<uint> SavedPalette { get; set; } = new();

    public List<ProtectedRange> ProtectedRanges { get; set; } = new();

    public List<ConditionalFormattingRule> ConditionalRules { get; set; } = new();

    public List<ActivityEntry> ActivityLog { get; set; } = new();

    public List<SheetVersionSnapshot> Snapshots { get; set; } = new();

    public List<SheetPresenceInfo> Presence { get; set; } = new();

    public List<SheetChatMessage> ChatMessages { get; set; } = new();

    public List<SheetPermissionPreset> PermissionPresets { get; set; } = new();

    public List<SheetMemberProfile> MemberProfiles { get; set; } = new();

    public List<SheetTabPreset> TabPresets { get; set; } = new();

    public List<SheetVersionSnapshot> SharedUndoHistory { get; set; } = new();

    public List<SheetVersionSnapshot> SharedRedoHistory { get; set; } = new();

    public List<TemporaryInviteCode> TemporaryInviteCodes { get; set; } = new();

    public List<InviteAuditEntry> InviteAuditLog { get; set; } = new();
}

public sealed class CellFeatureBundle
{
    public string CellKey { get; set; } = string.Empty;

    public List<CellComment> Comments { get; set; } = new();

    public DropdownValidation? Dropdown { get; set; }

    public bool IsChecklistCell { get; set; }

    public bool IsChecked { get; set; }

    public bool WrapText { get; set; }

    public CellVerticalAlign VerticalAlign { get; set; } = CellVerticalAlign.Middle;

    public CellBorderFlags Borders { get; set; }

    public bool Locked { get; set; }
}

public sealed class AdvancedSheetEnvelope
{
    public SheetDocument Document { get; set; } = new();

    public ExpandedSheetSettings Settings { get; set; } = new();

    public Dictionary<string, CellFeatureBundle> CellFeatures { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    [JsonIgnore]
    public Dictionary<string, string> Locks { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}
