using System.Text.Json;

namespace EZSheets.Models;

internal static class SheetSerializationHelper
{
    private static readonly JsonSerializerOptions JsonOptions = new();
    private const int MaxEmbeddedSnapshotDocumentChars = 1_500_000;

    public static void SanitizeLoadedDocumentInPlace(SheetDocument document)
    {
        document.Normalize(0, 0);
        document.Settings ??= new ExpandedSheetSettings();

        var settings = document.Settings;
        settings.Category ??= string.Empty;
        settings.ThemeName = string.IsNullOrWhiteSpace(settings.ThemeName) ? "Default" : settings.ThemeName;
        settings.SavedPalette ??= new List<uint>();
        settings.ProtectedRanges ??= new List<ProtectedRange>();
        settings.ConditionalRules ??= new List<ConditionalFormattingRule>();
        settings.ActivityLog ??= new List<ActivityEntry>();
        settings.Snapshots ??= new List<SheetVersionSnapshot>();
        settings.Presence ??= new List<SheetPresenceInfo>();
        settings.ChatMessages ??= new List<SheetChatMessage>();
        settings.PermissionPresets ??= new List<SheetPermissionPreset>();
        settings.MemberProfiles ??= new List<SheetMemberProfile>();
        settings.TabPresets ??= new List<SheetTabPreset>();
        settings.SharedUndoHistory ??= new List<SheetVersionSnapshot>();
        settings.SharedRedoHistory ??= new List<SheetVersionSnapshot>();
        settings.TemporaryInviteCodes ??= new List<TemporaryInviteCode>();
        settings.InviteAuditLog ??= new List<InviteAuditEntry>();

        TrimListTailInPlace(settings.ActivityLog, 180);
        TrimListTailInPlace(settings.Presence, 120);
        TrimListTailInPlace(settings.ChatMessages, 200);
        TrimListTailInPlace(settings.MemberProfiles, 120);
        TrimListTailInPlace(settings.TemporaryInviteCodes, 60);
        TrimListTailInPlace(settings.InviteAuditLog, 180);
        SanitizeSnapshotListInPlace(settings.Snapshots, 24);
        SanitizeSnapshotListInPlace(settings.SharedUndoHistory, 40);
        SanitizeSnapshotListInPlace(settings.SharedRedoHistory, 40);

        foreach (var tab in document.Tabs)
        {
            tab.Cells ??= new Dictionary<string, SheetCellData>(StringComparer.OrdinalIgnoreCase);
            tab.CellFeatures ??= new Dictionary<string, CellFeatureBundle>(StringComparer.OrdinalIgnoreCase);
            tab.ColumnWidths ??= new Dictionary<int, float>();
            tab.RowHeights ??= new Dictionary<int, float>();
        }
    }

    public static string SerializeForSnapshot(SheetDocument? document, int rowsCount, int colsCount)
    {
        var clone = CloneDocument(document, clearRuntimeState: true, trimForPersistence: false, aggressiveTrim: false, sanitizeSnapshotPayloads: true);
        clone.Normalize(rowsCount, colsCount);
        return JsonSerializer.Serialize(clone, JsonOptions);
    }

    public static SheetDocument CloneForPersistence(SheetDocument? document, int rowsCount, int colsCount, bool aggressiveTrim)
    {
        var clone = CloneDocument(document, clearRuntimeState: true, trimForPersistence: true, aggressiveTrim: aggressiveTrim, sanitizeSnapshotPayloads: true);
        clone.Normalize(rowsCount, colsCount);
        return clone;
    }

    public static SheetDocument CloneForLocalCache(SheetDocument? document, int rowsCount, int colsCount)
    {
        var clone = CloneDocument(document, clearRuntimeState: true, trimForPersistence: false, aggressiveTrim: false, sanitizeSnapshotPayloads: true);
        clone.Normalize(rowsCount, colsCount);
        return clone;
    }

    private static SheetDocument CloneDocument(SheetDocument? source, bool clearRuntimeState, bool trimForPersistence, bool aggressiveTrim, bool sanitizeSnapshotPayloads)
    {
        if (source is null)
        {
            return SheetDocument.CreateDefault();
        }

        var clone = new SheetDocument
        {
            Tabs = source.Tabs?.Select(CloneTab).ToList() ?? new List<SheetTabData>(),
            ActiveTabIndex = source.ActiveTabIndex,
            Settings = CloneSettings(source.Settings, clearRuntimeState, trimForPersistence, aggressiveTrim, sanitizeSnapshotPayloads),
            Cells = null,
        };

        clone.Normalize(0, 0);
        return clone;
    }

    private static SheetTabData CloneTab(SheetTabData source)
    {
        var tab = new SheetTabData
        {
            Name = source.Name ?? string.Empty,
            TabColor = source.TabColor,
            Cells = new Dictionary<string, SheetCellData>(StringComparer.OrdinalIgnoreCase),
            CellFeatures = new Dictionary<string, CellFeatureBundle>(StringComparer.OrdinalIgnoreCase),
            ColumnWidths = source.ColumnWidths is null ? new Dictionary<int, float>() : new Dictionary<int, float>(source.ColumnWidths),
            RowHeights = source.RowHeights is null ? new Dictionary<int, float>() : new Dictionary<int, float>(source.RowHeights),
        };

        if (source.Cells is not null)
        {
            foreach (var pair in source.Cells)
            {
                tab.Cells[pair.Key] = CloneCell(pair.Value);
            }
        }

        if (source.CellFeatures is not null)
        {
            foreach (var pair in source.CellFeatures)
            {
                tab.CellFeatures[pair.Key] = CloneFeature(pair.Value);
            }
        }

        return tab;
    }

    private static SheetCellData CloneCell(SheetCellData source)
    {
        return new SheetCellData
        {
            Value = source.Value ?? string.Empty,
            TextColor = source.TextColor,
            BackgroundColor = source.BackgroundColor,
            HorizontalAlign = source.HorizontalAlign ?? "left",
            Bold = source.Bold,
            Italic = source.Italic,
            Underline = source.Underline,
            FontScale = source.FontScale,
            WrapText = source.WrapText,
            VerticalAlign = source.VerticalAlign ?? "middle",
            Borders = source.Borders,
            BorderColor = source.BorderColor,
            Locked = source.Locked,
            MergedInto = source.MergedInto,
            RowSpan = source.RowSpan,
            ColSpan = source.ColSpan,
        };
    }

    private static CellFeatureBundle CloneFeature(CellFeatureBundle source)
    {
        return new CellFeatureBundle
        {
            CellKey = source.CellKey ?? string.Empty,
            Comments = source.Comments?.Select(CloneComment).ToList() ?? new List<CellComment>(),
            Dropdown = source.Dropdown is null ? null : CloneDropdown(source.Dropdown),
            IsChecklistCell = source.IsChecklistCell,
            IsChecked = source.IsChecked,
            WrapText = source.WrapText,
            VerticalAlign = source.VerticalAlign,
            Borders = source.Borders,
            Locked = source.Locked,
        };
    }

    private static CellComment CloneComment(CellComment source)
    {
        return new CellComment
        {
            Id = source.Id ?? string.Empty,
            AuthorId = source.AuthorId ?? string.Empty,
            AuthorName = source.AuthorName ?? string.Empty,
            Message = source.Message ?? string.Empty,
            CreatedAtUtc = source.CreatedAtUtc,
        };
    }

    private static DropdownValidation CloneDropdown(DropdownValidation source)
    {
        return new DropdownValidation
        {
            Options = source.Options is null ? new List<string>() : source.Options.Select(option => option ?? string.Empty).ToList(),
            AllowCustomText = source.AllowCustomText,
        };
    }

    private static ExpandedSheetSettings CloneSettings(ExpandedSheetSettings? source, bool clearRuntimeState, bool trimForPersistence, bool aggressiveTrim, bool sanitizeSnapshotPayloads)
    {
        source ??= new ExpandedSheetSettings();

        var activityMax = aggressiveTrim ? 80 : 180;
        var snapshotMax = aggressiveTrim ? 12 : 24;
        var sharedHistoryMax = aggressiveTrim ? 20 : 40;
        var temporaryCodeMax = aggressiveTrim ? 30 : 60;
        var inviteAuditMax = aggressiveTrim ? 90 : 180;
        var memberProfileMax = aggressiveTrim ? 64 : 120;

        var settings = new ExpandedSheetSettings
        {
            TemplateKind = source.TemplateKind,
            ViewMode = source.ViewMode,
            Category = source.Category ?? string.Empty,
            Favorite = source.Favorite,
            AutoSaveLocal = trimForPersistence ? false : source.AutoSaveLocal,
            AutoSaveCloud = trimForPersistence ? true : source.AutoSaveCloud,
            AutoSaveSeconds = trimForPersistence ? 20 : source.AutoSaveSeconds,
            FrozenRows = source.FrozenRows,
            FrozenColumns = source.FrozenColumns,
            ThemeName = source.ThemeName ?? "Default",
            SavedPalette = source.SavedPalette is null ? new List<uint>() : new List<uint>(source.SavedPalette),
            ProtectedRanges = source.ProtectedRanges?.Select(CloneProtectedRange).ToList() ?? new List<ProtectedRange>(),
            ConditionalRules = source.ConditionalRules?.Select(CloneConditionalRule).ToList() ?? new List<ConditionalFormattingRule>(),
            ActivityLog = TrimTail(source.ActivityLog?.Select(CloneActivityEntry).ToList(), activityMax),
            Snapshots = CloneSnapshotList(source.Snapshots, snapshotMax, sanitizeSnapshotPayloads),
            Presence = clearRuntimeState ? new List<SheetPresenceInfo>() : TrimTail(source.Presence?.Select(ClonePresenceInfo).ToList(), 120),
            ChatMessages = clearRuntimeState ? new List<SheetChatMessage>() : TrimTail(source.ChatMessages?.Select(CloneChatMessage).ToList(), 200),
            PermissionPresets = source.PermissionPresets?.Select(ClonePermissionPreset).ToList() ?? new List<SheetPermissionPreset>(),
            MemberProfiles = TrimTail(source.MemberProfiles?.Select(CloneMemberProfile).ToList(), memberProfileMax),
            TabPresets = source.TabPresets?.Select(CloneTabPreset).ToList() ?? new List<SheetTabPreset>(),
            SharedUndoHistory = CloneSnapshotList(source.SharedUndoHistory, sharedHistoryMax, sanitizeSnapshotPayloads),
            SharedRedoHistory = CloneSnapshotList(source.SharedRedoHistory, sharedHistoryMax, sanitizeSnapshotPayloads),
            TemporaryInviteCodes = TrimTail(source.TemporaryInviteCodes?.Select(CloneTemporaryInviteCode).ToList(), temporaryCodeMax),
            InviteAuditLog = TrimTail(source.InviteAuditLog?.Select(CloneInviteAuditEntry).ToList(), inviteAuditMax),
        };

        return settings;
    }

    private static ProtectedRange CloneProtectedRange(ProtectedRange source)
    {
        return new ProtectedRange
        {
            Id = source.Id ?? string.Empty,
            StartRow = source.StartRow,
            StartColumn = source.StartColumn,
            EndRow = source.EndRow,
            EndColumn = source.EndColumn,
            EditorUserIds = source.EditorUserIds is null ? new List<string>() : source.EditorUserIds.Select(x => x ?? string.Empty).ToList(),
            AllowOwnerOnly = source.AllowOwnerOnly,
        };
    }

    private static ConditionalFormattingRule CloneConditionalRule(ConditionalFormattingRule source)
    {
        return new ConditionalFormattingRule
        {
            Id = source.Id ?? string.Empty,
            Operator = source.Operator,
            TargetValue = source.TargetValue ?? string.Empty,
            TextColor = source.TextColor,
            BackgroundColor = source.BackgroundColor,
            Bold = source.Bold,
            Italic = source.Italic,
            Underline = source.Underline,
        };
    }

    private static ActivityEntry CloneActivityEntry(ActivityEntry source)
    {
        return new ActivityEntry
        {
            Id = source.Id ?? string.Empty,
            UserId = source.UserId ?? string.Empty,
            UserName = source.UserName ?? string.Empty,
            Action = source.Action ?? string.Empty,
            CellKey = source.CellKey,
            TabName = source.TabName,
            OldValue = source.OldValue,
            NewValue = source.NewValue,
            TimestampUtc = source.TimestampUtc,
        };
    }

    private static SheetVersionSnapshot CloneSnapshot(SheetVersionSnapshot source, bool sanitizeSnapshotPayloads)
    {
        return new SheetVersionSnapshot
        {
            Id = source.Id ?? string.Empty,
            Label = source.Label ?? "Snapshot",
            CreatedByUserId = source.CreatedByUserId ?? string.Empty,
            CreatedByName = source.CreatedByName ?? string.Empty,
            CreatedAtUtc = source.CreatedAtUtc,
            DocumentJson = sanitizeSnapshotPayloads ? CompactEmbeddedDocumentJson(source.DocumentJson) : (source.DocumentJson ?? string.Empty),
        };
    }

    private static List<SheetVersionSnapshot> CloneSnapshotList(List<SheetVersionSnapshot>? source, int maxCount, bool sanitizeSnapshotPayloads)
    {
        var cloned = source?.Select(entry => CloneSnapshot(entry, sanitizeSnapshotPayloads)).ToList() ?? new List<SheetVersionSnapshot>();
        return TrimTail(cloned, maxCount);
    }

    private static SheetPresenceInfo ClonePresenceInfo(SheetPresenceInfo source)
    {
        return new SheetPresenceInfo
        {
            UserId = source.UserId ?? string.Empty,
            UserName = source.UserName ?? string.Empty,
            ActiveTabName = source.ActiveTabName,
            EditingCellKey = source.EditingCellKey,
            LastSeenUtc = source.LastSeenUtc,
        };
    }

    private static SheetChatMessage CloneChatMessage(SheetChatMessage source)
    {
        return new SheetChatMessage
        {
            Id = source.Id ?? string.Empty,
            AuthorUserId = source.AuthorUserId ?? string.Empty,
            AuthorName = source.AuthorName ?? string.Empty,
            Message = source.Message ?? string.Empty,
            TimestampUtc = source.TimestampUtc,
        };
    }

    private static SheetPermissionPreset ClonePermissionPreset(SheetPermissionPreset source)
    {
        return new SheetPermissionPreset
        {
            Name = source.Name ?? "Viewer",
            Color = source.Color,
            Permissions = ClonePermissionSet(source.Permissions),
        };
    }

    private static SheetPermissionSet ClonePermissionSet(SheetPermissionSet? source)
    {
        source ??= new SheetPermissionSet();
        return new SheetPermissionSet
        {
            EditSheet = source.EditSheet,
            DeleteSheet = source.DeleteSheet,
            EditPermissions = source.EditPermissions,
            CreateTabs = source.CreateTabs,
            SeeHistory = source.SeeHistory,
            UseComments = source.UseComments,
            ImportSheet = source.ImportSheet,
            SaveLocal = source.SaveLocal,
            Invite = source.Invite,
            BlockUsers = source.BlockUsers,
            Admin = source.Admin,
        };
    }

    private static SheetMemberProfile CloneMemberProfile(SheetMemberProfile source)
    {
        return new SheetMemberProfile
        {
            UserId = source.UserId ?? string.Empty,
            CharacterName = source.CharacterName ?? string.Empty,
            JoinedAtUtc = source.JoinedAtUtc,
            LastSeenUtc = source.LastSeenUtc,
            AccessExpiresAtUtc = source.AccessExpiresAtUtc,
            AssignedPresetName = source.AssignedPresetName ?? "Viewer",
            RoleColor = source.RoleColor,
            IsBlocked = source.IsBlocked,
            Permissions = ClonePermissionSet(source.Permissions),
        };
    }

    private static SheetTabPreset CloneTabPreset(SheetTabPreset source)
    {
        return new SheetTabPreset
        {
            Name = source.Name ?? "Preset",
            Tab = CloneTab(source.Tab ?? new SheetTabData()),
        };
    }

    private static TemporaryInviteCode CloneTemporaryInviteCode(TemporaryInviteCode source)
    {
        return new TemporaryInviteCode
        {
            Id = source.Id ?? string.Empty,
            Code = source.Code ?? string.Empty,
            CreatedByUserId = source.CreatedByUserId ?? string.Empty,
            CreatedByName = source.CreatedByName ?? string.Empty,
            CreatedAtUtc = source.CreatedAtUtc,
            DurationMinutes = source.DurationMinutes,
            UsedAtUtc = source.UsedAtUtc,
            UsedByUserId = source.UsedByUserId ?? string.Empty,
            UsedByName = source.UsedByName ?? string.Empty,
            Invalidated = source.Invalidated,
        };
    }

    private static InviteAuditEntry CloneInviteAuditEntry(InviteAuditEntry source)
    {
        return new InviteAuditEntry
        {
            Id = source.Id ?? string.Empty,
            Code = source.Code ?? string.Empty,
            CodeType = source.CodeType ?? string.Empty,
            CreatedByUserId = source.CreatedByUserId ?? string.Empty,
            CreatedByName = source.CreatedByName ?? string.Empty,
            CreatedAtUtc = source.CreatedAtUtc,
            DurationMinutes = source.DurationMinutes,
            WasUsed = source.WasUsed,
            UsedAtUtc = source.UsedAtUtc,
            UsedByUserId = source.UsedByUserId ?? string.Empty,
            UsedByName = source.UsedByName ?? string.Empty,
        };
    }

    private static void SanitizeSnapshotListInPlace(List<SheetVersionSnapshot> snapshots, int maxCount)
    {
        TrimListTailInPlace(snapshots, maxCount);
        for (var index = 0; index < snapshots.Count; index++)
        {
            var snapshot = snapshots[index] ?? new SheetVersionSnapshot();
            snapshot.Id ??= string.Empty;
            snapshot.Label = string.IsNullOrWhiteSpace(snapshot.Label) ? "Snapshot" : snapshot.Label;
            snapshot.CreatedByUserId ??= string.Empty;
            snapshot.CreatedByName ??= string.Empty;
            snapshot.DocumentJson = NeedsEmbeddedDocumentCompaction(snapshot.DocumentJson)
                ? CompactEmbeddedDocumentJson(snapshot.DocumentJson)
                : snapshot.DocumentJson ?? string.Empty;
            snapshots[index] = snapshot;
        }
    }

    private static bool NeedsEmbeddedDocumentCompaction(string? documentJson)
    {
        if (string.IsNullOrWhiteSpace(documentJson) || documentJson.Length > MaxEmbeddedSnapshotDocumentChars)
        {
            return true;
        }

        return HasNonEmptyEmbeddedList(documentJson, "Snapshots")
            || HasNonEmptyEmbeddedList(documentJson, "SharedUndoHistory")
            || HasNonEmptyEmbeddedList(documentJson, "SharedRedoHistory")
            || HasNonEmptyEmbeddedList(documentJson, "Presence")
            || HasNonEmptyEmbeddedList(documentJson, "ChatMessages");
    }

    private static bool HasNonEmptyEmbeddedList(string documentJson, string propertyName)
    {
        var marker = $"\"{propertyName}\":[";
        if (!documentJson.Contains(marker, StringComparison.Ordinal))
        {
            return false;
        }

        var emptyMarker = $"\"{propertyName}\":[]";
        return !documentJson.Contains(emptyMarker, StringComparison.Ordinal);
    }

    private static void TrimListTailInPlace<T>(List<T> source, int maxCount)
    {
        if (source.Count <= maxCount)
        {
            return;
        }

        source.RemoveRange(0, source.Count - maxCount);
    }

    private static string CompactEmbeddedDocumentJson(string? documentJson)
    {
        if (string.IsNullOrWhiteSpace(documentJson) || documentJson.Length > MaxEmbeddedSnapshotDocumentChars)
        {
            return JsonSerializer.Serialize(SheetDocument.CreateDefault(), JsonOptions);
        }

        try
        {
            var document = JsonSerializer.Deserialize<SheetDocument>(documentJson, JsonOptions) ?? SheetDocument.CreateDefault();
            document.Normalize(0, 0);
            document.Settings ??= new ExpandedSheetSettings();
            document.Settings.Presence = new List<SheetPresenceInfo>();
            document.Settings.ChatMessages = new List<SheetChatMessage>();
            document.Settings.Snapshots = new List<SheetVersionSnapshot>();
            document.Settings.SharedUndoHistory = new List<SheetVersionSnapshot>();
            document.Settings.SharedRedoHistory = new List<SheetVersionSnapshot>();
            document.Settings.ActivityLog = TrimTail(document.Settings.ActivityLog?.Select(CloneActivityEntry).ToList(), 120);
            document.Settings.TemporaryInviteCodes = TrimTail(document.Settings.TemporaryInviteCodes?.Select(CloneTemporaryInviteCode).ToList(), 30);
            document.Settings.InviteAuditLog = TrimTail(document.Settings.InviteAuditLog?.Select(CloneInviteAuditEntry).ToList(), 90);
            document.Settings.MemberProfiles = TrimTail(document.Settings.MemberProfiles?.Select(CloneMemberProfile).ToList(), 64);
            return JsonSerializer.Serialize(document, JsonOptions);
        }
        catch
        {
            return JsonSerializer.Serialize(SheetDocument.CreateDefault(), JsonOptions);
        }
    }

    private static List<T> TrimTail<T>(List<T>? source, int maxCount)
    {
        if (source is null || source.Count == 0)
        {
            return new List<T>();
        }

        if (source.Count <= maxCount)
        {
            return source;
        }

        return source.Skip(Math.Max(0, source.Count - maxCount)).ToList();
    }
}
