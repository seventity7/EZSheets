using System.Globalization;
using System.Linq;
using System.Media;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Text.Json;
using Dalamud.Game.Gui.Toast;
using Dalamud.Interface.ImGuiNotification;
using Dalamud.Bindings.ImGui;
using EZSheets.Models;
using EZSheets.Services;

namespace EZSheets.Windows;

public sealed partial class MainWindow
{
    private readonly Stack<string> redoHistory = new();
    private bool advancedWindowOpen;
    private bool historyWindowOpen;
    private bool searchWindowOpen;
    private bool chatWindowOpen;
    private bool chatPinned = true;
    private string searchQuery = string.Empty;
    private readonly List<SearchResultItem> searchResults = new();
    private string chatInput = string.Empty;
    private string commentDraft = string.Empty;
    private string dropdownOptionsDraft = string.Empty;
    private string generalCategoryDraft = string.Empty;
    private int selectedTemplateIndex;
    private int selectedViewModeIndex;
    private int selectedOperatorIndex;
    private string ruleTargetValue = string.Empty;
    private bool ruleUseTextColor;
    private bool ruleUseBackgroundColor;
    private Vector4 ruleTextColor = new(1f, 1f, 1f, 1f);
    private Vector4 ruleBackgroundColor = new(0.2f, 0.2f, 0.2f, 1f);
    private Vector4 selectedBorderColor = new(1f, 1f, 1f, 0.8f);
    private bool permissionsWindowOpen;
    private string permissionsEditorUserId = string.Empty;
    private string permissionPresetDraft = string.Empty;
    private Vector4 permissionPresetColor = new(0.831f, 0.404f, 0.157f, 1f);
    private readonly List<SheetMemberRow> currentSheetMembers = new();
    private readonly List<SheetCellEditLock> currentCellLocks = new();
    private readonly List<SheetBlockedMemberRow> currentSheetBlocklist = new();
    private bool blocklistWindowOpen;
    private bool blockedJoinPopupOpen;
    private string blockedJoinMessage = "You are blocked from this sheet.";
    private string pendingRemovalReasonUserId = string.Empty;
    private string pendingRemovalReasonText = string.Empty;
    private bool ruleBold;
    private bool ruleItalic;
    private bool ruleUnderline;
    private DateTimeOffset lastAutoSaveTickUtc = DateTimeOffset.MinValue;
    private DateTimeOffset lastPresenceHeartbeatUtc = DateTimeOffset.MinValue;
    private DateTimeOffset lastLiveRefreshUtc = DateTimeOffset.MinValue;
    private DateTimeOffset lastSheetListRefreshUtc = DateTimeOffset.MinValue;
    private DateTimeOffset lastSheetMembersRefreshUtc = DateTimeOffset.MinValue;
    private DateTimeOffset lastBlocklistRefreshUtc = DateTimeOffset.MinValue;
    private bool sheetMembersRefreshInFlight;
    private bool blocklistRefreshInFlight;
    private bool liveRefreshInFlight;
    private bool removedByAdminPopupOpen;
    private string removedByAdminMessage = "You were removed by a admin.";
    private bool roleChangedPopupOpen;
    private string roleChangedOwnerName = "Sheet owner";
    private string roleChangedRoleName = "Viewer";
    private uint roleChangedRoleColor = 0xFF5C5C5Cu;
    private bool hasObservedRemoteRoleState;
    private bool sidebarChatScrollToBottomPending = true;
    private bool popupChatScrollToBottomPending = true;
    private string lastSidebarChatSheetId = string.Empty;
    private string lastPopupChatSheetId = string.Empty;
    private string lastObservedRemoteRoleSignature = string.Empty;
    private bool suppressNextCloudSaveHistory;
    private bool suppressNextCloudSaveStatus;
    private bool sheetListPinned = true;
    private bool sheetListWindowOpen;
    private bool presencePinned = true;
    private bool presenceWindowOpen;
    private string lastPersistedSheetFingerprint = string.Empty;
    private bool persistableSheetDirty;
    private DateTimeOffset lastPersistableChangeUtc = DateTimeOffset.MinValue;
    private int pendingBackgroundCloudSaveRequestId;
    private string lastObservedChatNotificationAnchor = string.Empty;
    private bool memberRemovedPopupOpen;
    private string memberRemovedName = "User";
    private readonly Dictionary<string, Vector2> pendingTabDeletePopupPositions = new(StringComparer.OrdinalIgnoreCase);
    private int selectedSheetPresetIndex;
    private string saveSheetPresetName = string.Empty;
    private string replaceWithQuery = string.Empty;
    private bool replaceAllTabs;
    private int unreadChatCount;
    private bool sharedHistoryOperationInFlight;
    private DateTimeOffset lastAutoReconnectAttemptUtc = DateTimeOffset.MinValue;
    private DateTimeOffset lastTemporaryAccessSweepUtc = DateTimeOffset.MinValue;
    private Vector2 lastPersistedChatWindowPos = new(float.NaN, float.NaN);
    private Vector2 lastPersistedChatWindowSize = new(float.NaN, float.NaN);
    private Vector2 lastPersistedSheetListWindowPos = new(float.NaN, float.NaN);
    private Vector2 lastPersistedSheetListWindowSize = new(float.NaN, float.NaN);
    private Vector2 lastPersistedPresenceWindowPos = new(float.NaN, float.NaN);
    private Vector2 lastPersistedPresenceWindowSize = new(float.NaN, float.NaN);
    private int temporaryInviteHoursSelection;
    private int temporaryInviteMinutesSelection;
    private string? cachedLatestTemporaryInviteCode;
    private bool autoSaveInFlight;
    private DateTimeOffset autoSaveBackoffUntilUtc = DateTimeOffset.MinValue;
    private bool chatHasUnreadAttention;
    private bool chatInputActive;
    private Vector2 mentionSuggestionAnchorMin;
    private Vector2 mentionSuggestionAnchorMax;
    private DateTimeOffset temporaryCodeCopiedFeedbackUntilUtc = DateTimeOffset.MinValue;
    private string lastMentionToastAnchor = string.Empty;
    private DateTimeOffset snapshotFeedbackUntilUtc = DateTimeOffset.MinValue;
    private bool homeViewActive;
    private Vector2 pendingSearchWindowPosition = new(float.NaN, float.NaN);
    private bool pendingSearchWindowReposition;

    private static readonly string[] TemplateNames = Enum.GetNames<SheetTemplateKind>();
    private static readonly string[] ViewModeNames = Enum.GetNames<SheetViewMode>();
    private static readonly string[] VerticalAlignNames = new[] { "Top", "Middle", "Bottom" };
    private static readonly string[] ConditionalOperatorNames = Enum.GetNames<ConditionalOperator>();

    [DllImport("user32.dll")]
    private static extern bool MessageBeep(uint type);

    [DllImport("winmm.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool PlaySound(string pszSound, nint hmod, uint fdwSound);

    private const uint SoundAlias = 0x00010000;
    private const uint SoundAsync = 0x0001;
    private const uint SoundNoDefault = 0x0002;

    private void InitializeAdvancedFeatures()
    {
        this.selectedTemplateIndex = 0;
        this.selectedViewModeIndex = 0;
        this.selectedOperatorIndex = 0;
        this.selectedBorderColor = ImGui.ColorConvertU32ToFloat4(this.configuration.LastBorderColor == 0 ? 0xCCFFFFFFu : this.configuration.LastBorderColor);
        this.permissionPresetColor = HexToVector4("#d46728");
        this.selectedSheetPresetIndex = 0;
    }

    private void UpdateAdvancedRuntime()
    {
        this.MaybeAutoReconnectLastSheet();

        if (this.currentSheet is null)
        {
            return;
        }

        this.EnsureAdvancedDefaults(this.currentSheet.Data);
        this.UpdatePresence();
        this.MaybeSyncPresenceHeartbeat();
        this.MaybeRefreshLiveSheetState();
        this.MaybeRefreshSheetList();
        this.MaybeEnforceTemporaryAccessExpiry();
        this.MaybeAutoSave();
    }

    private void DrawAdvancedQuickActionsRow()
    {
    }

    private void DrawAdvancedWindows()
    {
        this.DrawHistoryWindow();
        this.DrawSearchWindow();
        this.DrawChatWindow();
        this.DrawDetachedSheetListWindow();
        this.DrawDetachedPresenceWindow();
        this.DrawPermissionsWindow();
        this.DrawPermissionEditorWindow();
        this.DrawBlocklistWindow();
        this.DrawRemovedByAdminPopup();
        this.DrawBlockedJoinPopup();
        this.DrawRoleChangedPopup();
        this.DrawMemberRemovedPopup();
    }

    private void DrawTemplatePicker()
    {
        var templateIndex = this.selectedTemplateIndex;
        if (ImGui.Combo("Template", ref templateIndex, TemplateNames, TemplateNames.Length))
        {
            this.selectedTemplateIndex = templateIndex;
            var defaults = TemplateFactory.GetTemplateDefaults((SheetTemplateKind)this.selectedTemplateIndex);
            this.newRows = defaults.Rows;
            this.newCols = defaults.Columns;
            if (string.IsNullOrWhiteSpace(this.newSheetTitle) || string.Equals(this.newSheetTitle, "New Sheet", StringComparison.OrdinalIgnoreCase) || TemplateNames.Contains(this.newSheetTitle, StringComparer.OrdinalIgnoreCase))
            {
                this.newSheetTitle = defaults.DefaultTitle;
            }
        }
    }

    private SheetDocument CreateDocumentFromSelectedTemplate()
    {
        var kind = (SheetTemplateKind)this.selectedTemplateIndex;
        var envelope = TemplateFactory.CreateTemplate(kind);
        var defaults = TemplateFactory.GetTemplateDefaults(kind);
        envelope.Document.Settings.TemplateKind = kind;
        envelope.Document.Normalize(this.newRows > 0 ? this.newRows : defaults.Rows, this.newCols > 0 ? this.newCols : defaults.Columns);
        this.EnsureAdvancedDefaults(envelope.Document);
        return envelope.Document;
    }

    private void OnSheetLoaded(RemoteSheet remote)
    {
        SheetSerializationHelper.SanitizeLoadedDocumentInPlace(remote.Data);
        this.EnsureAdvancedDefaults(remote.Data);
        this.generalCategoryDraft = remote.Data.Settings.Category;
        this.selectedViewModeIndex = (int)remote.Data.Settings.ViewMode;
        this.hideSharedCode = false;
        this.pendingTabDeletePopupPositions.Clear();
        this.selectedSheetPresetIndex = 0;
        this.saveSheetPresetName = string.Empty;
        this.hasObservedRemoteRoleState = false;
        this.lastObservedRemoteRoleSignature = string.Empty;
        this.SyncCurrentUserMemberProfile(remote.Data, remote.CreatedAt);
        this.MarkPresence(remote.Data, "Opened sheet");
        this.lastLiveRefreshUtc = DateTimeOffset.MinValue;
        this.lastPresenceHeartbeatUtc = DateTimeOffset.MinValue;
        this.unreadChatCount = 0;
        remote.Data.Settings.ChatMessages = this.FilterVisibleChatMessages(remote.Data.Settings.ChatMessages)
            .OrderBy(x => x.TimestampUtc)
            .TakeLast(200)
            .ToList();
        this.SyncPersistedSheetFingerprint();
        this.SyncChatNotificationAnchor(remote.Data.Settings.ChatMessages);
    }

    private void OnSheetSavedToCloud(RemoteSheet saved)
    {
        SheetSerializationHelper.SanitizeLoadedDocumentInPlace(saved.Data);
        this.EnsureAdvancedDefaults(saved.Data);
        this.SyncCurrentUserMemberProfile(saved.Data, saved.CreatedAt);
        this.MarkPresence(saved.Data, "Saved to cloud");
        if (this.suppressNextCloudSaveHistory)
        {
            this.suppressNextCloudSaveHistory = false;
        }
        else
        {
            this.RecordActivity("Saved sheet to cloud", null, null, null);
        }
        this.SyncPersistedSheetFingerprint();
        this.SyncChatNotificationAnchor(saved.Data.Settings.ChatMessages);
        this.lastLiveRefreshUtc = DateTimeOffset.UtcNow;
    }

    private void EnsureAdvancedDefaults(SheetDocument document)
    {
        document.Settings ??= new ExpandedSheetSettings();
        this.EnsurePermissionDefaults(document.Settings);
        foreach (var tab in document.Tabs)
        {
            tab.CellFeatures ??= new Dictionary<string, CellFeatureBundle>(StringComparer.OrdinalIgnoreCase);
        }
    }

    private ExpandedSheetSettings GetSheetSettings()
    {
        if (this.currentSheet is null)
        {
            return new ExpandedSheetSettings();
        }

        this.EnsureAdvancedDefaults(this.currentSheet.Data);
        return this.currentSheet.Data.Settings;
    }

    private CellFeatureBundle? GetSelectedFeatureBundle(bool create)
    {
        var activeTab = this.GetActiveTab();
        if (activeTab is null)
        {
            return null;
        }

        var key = this.GetPrimarySelectedCellKey();
        return create ? activeTab.GetOrCreateFeature(key) : activeTab.TryGetFeature(key, out var feature) ? feature : null;
    }

    private void DrawHistoryWindow()
    {
        if (!this.historyWindowOpen || this.currentSheet is null)
        {
            return;
        }

        var open = this.historyWindowOpen;
        ImGui.SetNextWindowSize(new Vector2(620f, 420f), ImGuiCond.FirstUseEver);
        if (!ImGui.Begin("Sheet History###EZSheetsHistory", ref open))
        {
            ImGui.End();
            this.historyWindowOpen = open;
            return;
        }

        this.historyWindowOpen = open;
        if (ImGui.BeginTabBar("##sheet-history-tabs"))
        {
            if (ImGui.BeginTabItem("History"))
            {
                this.DrawAdvancedHistoryTab();
                ImGui.EndTabItem();
            }

            if (ImGui.BeginTabItem("Saves"))
            {
                this.DrawSaveHistoryTab();
                ImGui.EndTabItem();
            }

            if (ImGui.BeginTabItem("Cell"))
            {
                this.DrawCellHistoryTab();
                ImGui.EndTabItem();
            }

            if (this.currentRole == SheetAccessRole.Owner && ImGui.BeginTabItem("Invites"))
            {
                this.DrawInviteHistoryTab();
                ImGui.EndTabItem();
            }

            ImGui.EndTabBar();
        }
        ImGui.End();
    }

    private void DrawAdvancedWindow()
    {
        if (!this.advancedWindowOpen || this.currentSheet is null)
        {
            return;
        }

        var open = this.advancedWindowOpen;
        if (!ImGui.Begin("EZSheets Features###EZSheetsFeatures", ref open, ImGuiWindowFlags.AlwaysAutoResize))
        {
            ImGui.End();
            this.advancedWindowOpen = open;
            return;
        }

        this.advancedWindowOpen = open;
        if (ImGui.BeginTabBar("##features-tabs"))
        {
            if (ImGui.BeginTabItem("General"))
            {
                this.DrawAdvancedGeneralTab();
                ImGui.EndTabItem();
            }

            if (ImGui.BeginTabItem("Cell"))
            {
                this.DrawAdvancedCellTab();
                ImGui.EndTabItem();
            }

            if (ImGui.BeginTabItem("Rules"))
            {
                this.DrawAdvancedRulesTab();
                ImGui.EndTabItem();
            }

            if (ImGui.BeginTabItem("History"))
            {
                this.DrawAdvancedHistoryTab();
                ImGui.EndTabItem();
            }

            if (ImGui.BeginTabItem("Presence"))
            {
                this.DrawAdvancedPresenceTab();
                ImGui.EndTabItem();
            }

            ImGui.EndTabBar();
        }

        ImGui.End();
    }

    private void DrawAdvancedGeneralTab()
    {
        var settings = this.GetSheetSettings();
        var favorite = settings.Favorite;
        if (ImGui.Checkbox("Favorite this sheet", ref favorite))
        {
            settings.Favorite = favorite;
            this.RecordActivity("Sheet favorite changed", null, null, favorite ? "true" : "false");
        }

        this.generalCategoryDraft = string.IsNullOrWhiteSpace(this.generalCategoryDraft) ? settings.Category : this.generalCategoryDraft;
        if (ImGui.InputText("Category", ref this.generalCategoryDraft, 64))
        {
            settings.Category = this.generalCategoryDraft.Trim();
        }

        var viewMode = (int)settings.ViewMode;
        if (ImGui.Combo("View mode", ref viewMode, ViewModeNames, ViewModeNames.Length))
        {
            settings.ViewMode = (SheetViewMode)viewMode;
        }

        settings.AutoSaveLocal = false;
        settings.AutoSaveCloud = true;
        settings.AutoSaveSeconds = 20;
        ImGui.TextDisabled("Cloud auto-save is always enabled.");

        var frozenRows = settings.FrozenRows;
        if (ImGui.SliderInt("Frozen rows", ref frozenRows, 0, 5))
        {
            settings.FrozenRows = frozenRows;
        }

        var frozenCols = settings.FrozenColumns;
        if (ImGui.SliderInt("Frozen columns", ref frozenCols, 0, 3))
        {
            settings.FrozenColumns = frozenCols;
        }

        var themeName = settings.ThemeName;
        if (ImGui.InputText("Theme name", ref themeName, 64))
        {
            settings.ThemeName = themeName;
        }

        if (ImGui.Button("Create snapshot"))
        {
            this.CreateSnapshot();
        }

        ImGui.SameLine();
        if (ImGui.Button("Duplicate active tab"))
        {
            this.DuplicateActiveTab();
        }

        ImGui.SameLine();
        if (ImGui.Button("Save current text color to palette"))
        {
            var activeTab = this.GetActiveTab();
            if (activeTab is not null)
            {
                var cell = activeTab.GetOrCreateCell(this.GetPrimarySelectedCellKey());
                if (!settings.SavedPalette.Contains(cell.TextColor))
                {
                    settings.SavedPalette.Add(cell.TextColor);
                }
            }
        }

        if (settings.SavedPalette.Count > 0)
        {
            ImGui.TextUnformatted("Saved palette:");
            for (var i = 0; i < settings.SavedPalette.Count; i++)
            {
                if (i > 0)
                {
                    ImGui.SameLine();
                }

                var color = ImGui.ColorConvertU32ToFloat4(settings.SavedPalette[i]);
                if (this.DrawColorSwatchButton($"palette-{i}", color, new Vector2(16f, 16f)))
                {
                    this.PushUndoSnapshot();
                    this.ApplyToSelectedCells(cell => cell.TextColor = settings.SavedPalette[i]);
                }
            }
        }
    }

    private void DrawAdvancedCellTab()
    {
        var activeTab = this.GetActiveTab();
        if (activeTab is null)
        {
            return;
        }

        var key = this.GetPrimarySelectedCellKey();
        var cell = activeTab.GetOrCreateCell(key);
        var feature = activeTab.GetOrCreateFeature(key);
        ImGui.TextDisabled($"Cell: {ToA1(key)}");

        var wrap = cell.WrapText;
        if (ImGui.Checkbox("Wrap text", ref wrap))
        {
            this.PushUndoSnapshot();
            this.ApplyToSelectedCells(c => c.WrapText = wrap);
        }

        var valign = cell.VerticalAlign switch
        {
            "top" => 0,
            "bottom" => 2,
            _ => 1,
        };
        if (ImGui.Combo("Vertical align", ref valign, VerticalAlignNames, VerticalAlignNames.Length))
        {
            var value = valign == 0 ? "top" : valign == 2 ? "bottom" : "middle";
            this.PushUndoSnapshot();
            this.ApplyToSelectedCells(c => c.VerticalAlign = value);
        }

        var borderLeft = cell.Borders.HasFlag(CellBorderFlags.Left);
        var borderTop = cell.Borders.HasFlag(CellBorderFlags.Top);
        var borderRight = cell.Borders.HasFlag(CellBorderFlags.Right);
        var borderBottom = cell.Borders.HasFlag(CellBorderFlags.Bottom);
        ImGui.TextUnformatted("Borders");
        var borderChanged = false;
        borderChanged |= ImGui.Checkbox("Left", ref borderLeft);
        borderChanged |= ImGui.Checkbox("Top", ref borderTop);
        borderChanged |= ImGui.Checkbox("Right", ref borderRight);
        borderChanged |= ImGui.Checkbox("Bottom", ref borderBottom);
        if (borderChanged)
        {
            var flags = CellBorderFlags.None;
            if (borderLeft) flags |= CellBorderFlags.Left;
            if (borderTop) flags |= CellBorderFlags.Top;
            if (borderRight) flags |= CellBorderFlags.Right;
            if (borderBottom) flags |= CellBorderFlags.Bottom;
            this.PushUndoSnapshot();
            this.ApplyToSelectedCells(c => c.Borders = flags);
        }

        var checklist = feature.IsChecklistCell;
        if (ImGui.Checkbox("Checklist cell", ref checklist))
        {
            this.PushUndoSnapshot();
            foreach (var selected in this.selectedCellKeys)
            {
                activeTab.GetOrCreateFeature(selected).IsChecklistCell = checklist;
            }
        }

        if (feature.IsChecklistCell)
        {
            var isChecked = feature.IsChecked;
            if (ImGui.Checkbox("Checked", ref isChecked))
            {
                feature.IsChecked = isChecked;
            }
        }

        this.dropdownOptionsDraft = feature.Dropdown is null ? this.dropdownOptionsDraft : string.Join(", ", feature.Dropdown.Options);
        if (ImGui.InputText("Dropdown options", ref this.dropdownOptionsDraft, 512))
        {
        }

        if (ImGui.Button("Apply dropdown to selection"))
        {
            this.PushUndoSnapshot();
            var options = this.dropdownOptionsDraft.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).ToList();
            foreach (var selected in this.selectedCellKeys)
            {
                activeTab.GetOrCreateFeature(selected).Dropdown = new DropdownValidation { Options = options, AllowCustomText = true };
            }
        }

        ImGui.Separator();
        ImGui.TextUnformatted("Comments");
        foreach (var comment in feature.Comments.OrderByDescending(x => x.CreatedAtUtc).Take(10))
        {
            ImGui.TextDisabled($"{comment.AuthorName} • {comment.CreatedAtUtc.LocalDateTime:g}");
            ImGui.TextWrapped(comment.Message);
            ImGui.Separator();
        }

        ImGui.InputTextMultiline("##comment-draft", ref this.commentDraft, 1024, new System.Numerics.Vector2(360f, 70f));
        if (ImGui.Button("Add comment"))
        {
            if (!string.IsNullOrWhiteSpace(this.commentDraft))
            {
                feature.Comments.Add(new CellComment
                {
                    AuthorId = this.configuration.UserId,
                    AuthorName = this.GetSafeCharacterFullName(),
                    Message = this.commentDraft.Trim(),
                    CreatedAtUtc = DateTimeOffset.UtcNow,
                });
                this.RecordActivity("Comment added", key, null, this.commentDraft.Trim());
                this.commentDraft = string.Empty;
            }
        }

        ImGui.SameLine();
        if (ImGui.Button("Protect selection range"))
        {
            this.AddProtectedRangeFromSelection();
        }

        ImGui.SameLine();
        if (ImGui.Button("Remove protection"))
        {
            this.RemoveProtectionFromSelection();
        }
    }

    private void DrawAdvancedRulesTab()
    {
        var settings = this.GetSheetSettings();
        ImGui.TextWrapped("Conditional formatting rules are applied to every cell based on its displayed value.");
        ImGui.InputText("Target value", ref this.ruleTargetValue, 128);
        ImGui.Combo("Operator", ref this.selectedOperatorIndex, ConditionalOperatorNames, ConditionalOperatorNames.Length);
        ImGui.Checkbox("Use text color", ref this.ruleUseTextColor);
        if (this.ruleUseTextColor)
        {
            ImGui.ColorEdit4("Rule text color", ref this.ruleTextColor, ImGuiColorEditFlags.NoInputs);
        }
        ImGui.Checkbox("Use fill color", ref this.ruleUseBackgroundColor);
        if (this.ruleUseBackgroundColor)
        {
            ImGui.ColorEdit4("Rule fill color", ref this.ruleBackgroundColor, ImGuiColorEditFlags.NoInputs);
        }
        ImGui.Checkbox("Bold", ref this.ruleBold);
        ImGui.Checkbox("Italic", ref this.ruleItalic);
        ImGui.Checkbox("Underline", ref this.ruleUnderline);

        if (ImGui.Button("Add rule"))
        {
            settings.ConditionalRules.Add(new ConditionalFormattingRule
            {
                Operator = (ConditionalOperator)this.selectedOperatorIndex,
                TargetValue = this.ruleTargetValue.Trim(),
                TextColor = this.ruleUseTextColor ? ImGui.ColorConvertFloat4ToU32(this.ruleTextColor) : null,
                BackgroundColor = this.ruleUseBackgroundColor ? ImGui.ColorConvertFloat4ToU32(this.ruleBackgroundColor) : null,
                Bold = this.ruleBold ? true : null,
                Italic = this.ruleItalic ? true : null,
                Underline = this.ruleUnderline ? true : null,
            });
            this.ruleTargetValue = string.Empty;
        }

        ImGui.Separator();
        for (var i = 0; i < settings.ConditionalRules.Count; i++)
        {
            var rule = settings.ConditionalRules[i];
            ImGui.TextUnformatted($"{rule.Operator} {rule.TargetValue}");
            ImGui.SameLine();
            if (ImGui.SmallButton($"Delete##rule-{i}"))
            {
                settings.ConditionalRules.RemoveAt(i);
                i--;
            }
        }
    }

    private void DrawAdvancedHistoryTab()
    {
        var settings = this.GetSheetSettings();
        if (ImGui.Button("Create snapshot now"))
        {
            this.CreateSnapshot();
        }

        ImGui.Separator();
        ImGui.TextUnformatted("Snapshots");
        foreach (var snapshot in settings.Snapshots.OrderByDescending(x => x.CreatedAtUtc).Take(15).ToList())
        {
            ImGui.TextDisabled($"{snapshot.Label} • {snapshot.CreatedByName} • {snapshot.CreatedAtUtc.LocalDateTime:g}");
            ImGui.SameLine();
            if (ImGui.SmallButton($"Restore##{snapshot.Id}"))
            {
                this.RestoreSnapshot(snapshot.Id);
            }
            ImGui.SameLine();
            if (ImGui.SmallButton($"Delete##snap-{snapshot.Id}"))
            {
                settings.Snapshots.RemoveAll(x => x.Id == snapshot.Id);
            }
        }

        ImGui.Separator();
        ImGui.TextUnformatted("Recent activity");
        foreach (var entry in settings.ActivityLog.OrderByDescending(x => x.TimestampUtc).Where(entry => !this.IsSaveActivity(entry) && this.CanSeeHistoryEntry(entry)).Take(50))
        {
            this.DrawHistoryEntry(entry);
        }
    }

    private void DrawSaveHistoryTab()
    {
        var settings = this.GetSheetSettings();
        foreach (var entry in settings.ActivityLog.OrderByDescending(x => x.TimestampUtc).Where(entry => this.IsSaveActivity(entry) && this.CanSeeHistoryEntry(entry)).Take(80))
        {
            this.DrawHistoryEntry(entry);
        }
    }

    private void DrawCellHistoryTab()
    {
        var settings = this.GetSheetSettings();
        var cellKey = this.GetPrimarySelectedCellKey();
        ImGui.TextDisabled($"Selected cell: {ToA1(cellKey)}");
        ImGui.Separator();
        foreach (var entry in settings.ActivityLog
            .OrderByDescending(x => x.TimestampUtc)
            .Where(x => string.Equals(x.CellKey, cellKey, StringComparison.OrdinalIgnoreCase) && this.CanSeeHistoryEntry(x))
            .Take(40))
        {
            this.DrawHistoryEntry(entry);
        }
    }


    private void DrawInviteHistoryTab()
    {
        var settings = this.GetSheetSettings();
        var temporaryCodes = settings.TemporaryInviteCodes
            .Select(x => x.Code)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var inviteEntries = settings.InviteAuditLog
            .Where(x => !(string.Equals(x.CodeType, "unique", StringComparison.OrdinalIgnoreCase) && temporaryCodes.Contains(x.Code)))
            .OrderByDescending(x => x.CreatedAtUtc)
            .ThenByDescending(x => x.UsedAtUtc)
            .Take(150)
            .ToList();

        if (inviteEntries.Count == 0)
        {
            ImGui.TextDisabled("No invite codes have been recorded yet.");
            return;
        }

        foreach (var entry in inviteEntries)
        {
            var typeLabel = string.IsNullOrWhiteSpace(entry.CodeType) ? "Unknown" : CultureInfo.InvariantCulture.TextInfo.ToTitleCase(entry.CodeType.Trim());
            var createdBy = string.IsNullOrWhiteSpace(entry.CreatedByName) ? "Unknown" : entry.CreatedByName;
            ImGui.TextWrapped($"{typeLabel} • {entry.Code} • created by {createdBy} • {entry.CreatedAtUtc.LocalDateTime:g}");
            if (entry.DurationMinutes > 0)
            {
                ImGui.TextDisabled($"Duration: {this.FormatDurationMinutes(entry.DurationMinutes)}");
            }

            if (entry.WasUsed)
            {
                var usedBy = string.IsNullOrWhiteSpace(entry.UsedByName) ? "Unknown" : entry.UsedByName;
                ImGui.TextDisabled($"Used by: {usedBy} • {entry.UsedAtUtc?.LocalDateTime:g}");
            }
            else
            {
                ImGui.TextDisabled("Used by: not used yet");
            }

            ImGui.Spacing();
        }
    }

    private string FormatDurationMinutes(int totalMinutes)
    {
        totalMinutes = Math.Max(0, totalMinutes);
        var hours = totalMinutes / 60;
        var minutes = totalMinutes % 60;
        if (hours > 0 && minutes > 0)
        {
            return $"{hours:00}:{minutes:00}";
        }

        if (hours > 0)
        {
            return $"{hours:00}:00";
        }

        return $"00:{minutes:00}";
    }

    private void DrawHistoryEntry(ActivityEntry entry)
    {
        ImGui.TextWrapped($"{entry.TimestampUtc.LocalDateTime:t} • {entry.UserName} • {entry.Action}{(string.IsNullOrWhiteSpace(entry.CellKey) ? string.Empty : $" ({ToA1(entry.CellKey!)})")}");
        if (!string.IsNullOrWhiteSpace(entry.OldValue) || !string.IsNullOrWhiteSpace(entry.NewValue))
        {
            ImGui.TextDisabled($"Before: {entry.OldValue ?? "<empty>"}");
            ImGui.TextDisabled($"After: {entry.NewValue ?? "<empty>"}");
        }
        ImGui.Spacing();
    }

    private bool IsSaveActivity(ActivityEntry entry)
    {
        var action = entry.Action ?? string.Empty;
        return action.Contains("save", StringComparison.OrdinalIgnoreCase)
               || action.Contains("import", StringComparison.OrdinalIgnoreCase)
               || action.Contains("autosave", StringComparison.OrdinalIgnoreCase)
               || action.Contains("auto-save", StringComparison.OrdinalIgnoreCase);
    }

    private bool CanSeeHistoryEntry(ActivityEntry entry)
    {
        if (this.currentRole == SheetAccessRole.Owner)
        {
            return true;
        }

        return !(entry.Action ?? string.Empty).Contains("Generated a unique code", StringComparison.OrdinalIgnoreCase);
    }

    private void DrawAdvancedPresenceTab()
    {
        var settings = this.GetSheetSettings();
        ImGui.TextUnformatted("Users seen in this sheet");
        this.DrawPresenceEntries(settings);
    }

    private void DrawSearchWindow()
    {
        if (!this.searchWindowOpen || this.currentSheet is null)
        {
            return;
        }

        var open = this.searchWindowOpen;
        if (!ImGui.Begin("Find in Sheet###EZSheetsFind", ref open, ImGuiWindowFlags.AlwaysAutoResize))
        {
            ImGui.End();
            this.searchWindowOpen = open;
            return;
        }

        this.searchWindowOpen = open;
        if (ImGui.InputText("Search", ref this.searchQuery, 128))
        {
            this.BuildSearchResults();
        }

        if (ImGui.Button("Refresh results"))
        {
            this.BuildSearchResults();
        }

        ImGui.SameLine(0f, 6f);
        ImGui.SetNextItemWidth(160f);
        ImGui.InputText("Replace with", ref this.replaceWithQuery, 128);
        ImGui.SameLine(0f, 6f);
        ImGui.Checkbox("All tabs", ref this.replaceAllTabs);
        ImGui.SameLine(0f, 6f);
        if (ImGui.Button("Replace all") && !string.IsNullOrWhiteSpace(this.searchQuery))
        {
            this.ReplaceAllMatches();
        }

        ImGui.Separator();
        foreach (var result in this.searchResults.Take(100))
        {
            if (ImGui.Selectable($"{result.TabName} • {result.CellLabel} • {result.Preview}"))
            {
                var document = this.currentSheet.Data;
                document.ActiveTabIndex = result.TabIndex;
                this.SelectSingleCell(result.CellKey);
                this.selectedCellKey = result.CellKey;
            }
        }

        ImGui.End();
    }

    private void DrawChatWindow()
    {
        if (!this.chatWindowOpen || this.currentSheet is null || this.chatPinned)
        {
            return;
        }

        var open = this.chatWindowOpen;
        this.TryRestoreDetachedWindowLayout("chat");
        ImGui.SetNextWindowSizeConstraints(new Vector2(500f, 380f), new Vector2(1600f, 1200f));
        ImGui.SetNextWindowSize(new Vector2(540f, 430f), ImGuiCond.FirstUseEver);
        var chatWindowTitle = this.unreadChatCount > 0 ? $"Sheet Chat ({this.unreadChatCount})###EZSheetsChat" : "Sheet Chat###EZSheetsChat";
        if (!ImGui.Begin(chatWindowTitle, ref open))
        {
            ImGui.End();
            this.chatWindowOpen = open;
            return;
        }

        this.chatWindowOpen = open;
        this.TryPersistDetachedWindowLayout("chat");
        var settings = this.GetSheetSettings();
        var popupChatMessages = settings.ChatMessages.OrderBy(x => x.TimestampUtc).TakeLast(100).ToList();
        var popupSheetId = this.currentSheet.Id ?? string.Empty;
        if (!string.Equals(this.lastPopupChatSheetId, popupSheetId, StringComparison.OrdinalIgnoreCase))
        {
            this.lastPopupChatSheetId = popupSheetId;
            this.popupChatScrollToBottomPending = true;
        }

        var headerStart = ImGui.GetCursorPos();
        var headerWidth = ImGui.GetContentRegionAvail().X;
        var pinSize = new Vector2(22f, 22f);
        if (this.DrawStyledButton("chat-pin-window", "♯", HexToVector4("#4b4b4b"), WhiteText, pinSize))
        {
            this.chatPinned = true;
            this.chatWindowOpen = false;
            this.sidebarChatScrollToBottomPending = true;
        }
        if (ImGui.IsItemHovered())
        {
            ImGui.SetTooltip("Pin chat");
        }

        ImGui.SameLine(0f, 6f);
        var popupHeaderTitle = this.unreadChatCount > 0 ? $"Sheet Live Chat ({this.unreadChatCount})" : "Sheet Live Chat";
        this.DrawFakeBoldText(popupHeaderTitle, WhiteText);
        var localTime = DateTime.Now;
        var use24Hour = !CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern.Contains("tt", StringComparison.OrdinalIgnoreCase);
        var clockText = localTime.ToString(use24Hour ? "HH:mm:ss" : "hh:mm:ss", CultureInfo.CurrentCulture);
        var clockSize = ImGui.CalcTextSize(clockText);
        var clockX = headerStart.X + MathF.Max(0f, headerWidth - clockSize.X);
        var minClockX = headerStart.X + pinSize.X + 6f + ImGui.CalcTextSize(popupHeaderTitle).X + 8f;
        if (clockX > minClockX)
        {
            ImGui.SetCursorPos(new Vector2(clockX, headerStart.Y + 2f));
            ImGui.TextDisabled(clockText);
            if (ImGui.IsItemHovered())
            {
                ImGui.SetTooltip("Local Time");
            }
        }

        ImGui.Spacing();
        var inputHeight = 92f;
        var sendHeight = 28f;
        var childHeight = MathF.Max(220f, ImGui.GetContentRegionAvail().Y - inputHeight - sendHeight - 14f);

        ImGui.PushStyleColor(ImGuiCol.ChildBg, this.GetChatPanelBackgroundColor());
        ImGui.BeginChild("##chat-scroll", new Vector2(-1f, childHeight), true);
        foreach (var message in popupChatMessages)
        {
            this.DrawChatMessageEntry(message, settings);
        }
        ImGui.Dummy(new Vector2(1f, 1f));
        if (this.popupChatScrollToBottomPending || ImGui.IsWindowAppearing())
        {
            ImGui.SetScrollHereY(1f);
            this.popupChatScrollToBottomPending = false;
        }
        if (ImGui.IsWindowHovered() && ImGui.IsMouseClicked(ImGuiMouseButton.Left))
        {
            this.ClearUnreadChatAttention();
        }
        ImGui.EndChild();
        ImGui.PopStyleColor();

        this.DrawPromptedChatInputMultiline("##chat-input-frame", "##chat-input", ref this.chatInput, 1024, new Vector2(-1f, inputHeight));
        this.DrawChatMentionSuggestions("popup", ref this.chatInput);
        if (this.DrawStyledButton("chat-send-window", "Send message", AccentBlue, WhiteText, new Vector2(118f, 24f)))
        {
            this.TrySendChatMessage(settings);
        }

        ImGui.End();
    }


    private void DrawSidebarPresencePanel()
    {
        if (this.currentSheet is null || !this.presencePinned)
        {
            return;
        }

        this.DrawSectionDivider();
        var settings = this.GetSheetSettings();
        var height = 88f;
        var start = ImGui.GetCursorPos();
        this.DrawFakeBoldText("Users seen in this sheet", WhiteText);
        var btnSize = new Vector2(20f, 20f);
        var rightX = start.X + MathF.Max(0f, ImGui.GetContentRegionAvail().X - btnSize.X);
        ImGui.SetCursorPos(new Vector2(rightX, start.Y - 1f));
        if (this.DrawStyledButton("presence-pop-pin", "♯", HexToVector4("#3b3b3b"), WhiteText, btnSize))
        {
            this.presencePinned = false;
            this.presenceWindowOpen = true;
        }
        if (ImGui.IsItemHovered())
        {
            ImGui.SetTooltip("Pop out users panel");
        }

        ImGui.PushStyleColor(ImGuiCol.ChildBg, new Vector4(0.18f, 0.18f, 0.18f, 0.42f));
        if (ImGui.BeginChild("##presence-sidebar", new Vector2(-1f, height), true, ImGuiWindowFlags.AlwaysVerticalScrollbar))
        {
            this.DrawPresenceEntries(settings);

            if (ImGui.IsWindowHovered() && ImGui.IsMouseClicked(ImGuiMouseButton.Left))
            {
                this.ClearUnreadChatAttention();
            }

            var min = ImGui.GetWindowPos();
            var max = min + ImGui.GetWindowSize();
            var drawList = ImGui.GetWindowDrawList();
            drawList.AddRectFilledMultiColor(
                new Vector2(min.X, max.Y - 24f),
                max,
                ImGui.ColorConvertFloat4ToU32(new Vector4(0f, 0f, 0f, 0f)),
                ImGui.ColorConvertFloat4ToU32(new Vector4(0f, 0f, 0f, 0f)),
                ImGui.ColorConvertFloat4ToU32(new Vector4(0.08f, 0.08f, 0.08f, 0.65f)),
                ImGui.ColorConvertFloat4ToU32(new Vector4(0.08f, 0.08f, 0.08f, 0.65f)));
        }
        ImGui.EndChild();
        ImGui.PopStyleColor();
    }

    private void DrawDetachedPresenceWindow()
    {
        if (!this.presenceWindowOpen || this.currentSheet is null)
        {
            return;
        }

        var open = this.presenceWindowOpen;
        this.TryRestoreDetachedWindowLayout("presence");
        ImGui.SetNextWindowSize(new Vector2(420f, 250f), ImGuiCond.FirstUseEver);
        if (!ImGui.Begin("Users seen in this sheet###EZSheetsPresence", ref open))
        {
            ImGui.End();
            this.presenceWindowOpen = open;
            if (!open)
            {
                this.presencePinned = true;
            }
            return;
        }

        this.presenceWindowOpen = open;
        this.TryPersistDetachedWindowLayout("presence");
        var btnSize = new Vector2(20f, 20f);
        var rightX = MathF.Max(0f, ImGui.GetContentRegionAvail().X - btnSize.X);
        ImGui.SetCursorPos(new Vector2(rightX, ImGui.GetCursorPosY()));
        if (this.DrawStyledButton("presence-pin-window", "♯", HexToVector4("#4b4b4b"), WhiteText, btnSize))
        {
            this.presencePinned = true;
            this.presenceWindowOpen = false;
        }
        if (ImGui.IsItemHovered())
        {
            ImGui.SetTooltip("Pin users panel");
        }

        var settings = this.GetSheetSettings();
        ImGui.PushStyleColor(ImGuiCol.ChildBg, new Vector4(0.18f, 0.18f, 0.18f, 0.42f));
        if (ImGui.BeginChild("##presence-window-scroll", new Vector2(-1f, -1f), true, ImGuiWindowFlags.AlwaysVerticalScrollbar))
        {
            this.DrawPresenceEntries(settings);
        }
        ImGui.EndChild();
        ImGui.PopStyleColor();
        ImGui.End();
        if (!open)
        {
            this.presencePinned = true;
        }
    }

    private void DrawPresenceEntries(ExpandedSheetSettings settings)
    {
        foreach (var presence in settings.Presence.OrderByDescending(x => x.LastSeenUtc).Take(30).ToList())
        {
            var color = this.GetUserAccentColor(presence.UserId);
            var status = this.GetPresenceStatusLabel(presence);
            ImGui.TextColored(color, "■");
            ImGui.SameLine(0f, 6f);
            ImGui.TextColored(color, string.IsNullOrWhiteSpace(presence.UserName) ? "Unknown user" : presence.UserName);
            ImGui.SameLine(0f, 6f);
            ImGui.TextDisabled($"· {status} · Tab: {presence.ActiveTabName ?? "-"} · Cell: {presence.EditingCellKey ?? "-"}");
            ImGui.Spacing();
        }
    }

    private void DrawSidebarActivityFeedPanel()
    {
        if (this.currentSheet is null)
        {
            return;
        }

        this.DrawSectionDivider();
        this.DrawFakeBoldText("Activity Feed", WhiteText);
        var settings = this.GetSheetSettings();
        ImGui.PushStyleColor(ImGuiCol.ChildBg, new Vector4(0.18f, 0.18f, 0.18f, 0.42f));
        if (ImGui.BeginChild("##activity-feed-sidebar", new Vector2(-1f, 98f), true, ImGuiWindowFlags.AlwaysVerticalScrollbar))
        {
            foreach (var entry in settings.ActivityLog.OrderByDescending(x => x.TimestampUtc).Where(x => !this.IsSaveActivity(x)).Take(8))
            {
                ImGui.TextColored(this.GetUserAccentColor(entry.UserId), string.IsNullOrWhiteSpace(entry.UserName) ? "User" : entry.UserName);
                ImGui.SameLine(0f, 4f);
                ImGui.TextDisabled($"· {entry.Action}");
                if (!string.IsNullOrWhiteSpace(entry.CellKey))
                {
                    ImGui.SameLine(0f, 4f);
                    ImGui.TextDisabled($"({ToA1(entry.CellKey)})");
                }
                ImGui.Spacing();
            }
        }
        ImGui.EndChild();
        ImGui.PopStyleColor();
    }

    private void DrawSidebarChatPanel()
    {
        if (this.currentSheet is null || !this.supabase.IsAuthenticated || !this.chatPinned)
        {
            return;
        }

        this.DrawSectionDivider();
        var start = ImGui.GetCursorPos();
        var availableWidth = ImGui.GetContentRegionAvail().X;
        var chatTitle = this.unreadChatCount > 0 ? $"Sheet Live Chat ({this.unreadChatCount})" : "Sheet Live Chat";
        this.DrawFakeBoldText(chatTitle, WhiteText);
        var btnSize = new Vector2(20f, 20f);
        var rightX = start.X + MathF.Max(0f, availableWidth - btnSize.X);
        var localTime = DateTime.Now;
        var use24Hour = !CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern.Contains("tt", StringComparison.OrdinalIgnoreCase);
        var clockText = localTime.ToString(use24Hour ? "HH:mm:ss" : "hh:mm:ss", CultureInfo.CurrentCulture);
        var titleWidth = ImGui.CalcTextSize(chatTitle).X;
        var clockSize = ImGui.CalcTextSize(clockText);
        var desiredClockX = start.X + titleWidth + 8f;
        var maxClockX = rightX - clockSize.X - 8f;
        if (maxClockX > desiredClockX)
        {
            ImGui.SetCursorPos(new Vector2(MathF.Min(desiredClockX, maxClockX), start.Y));
            ImGui.TextDisabled(clockText);
            if (ImGui.IsItemHovered())
            {
                ImGui.SetTooltip("Local Time");
            }
        }
        ImGui.SetCursorPos(new Vector2(rightX, start.Y - 1f));
        if (this.DrawStyledButton("chat-pop-pin", "♯", HexToVector4("#3b3b3b"), WhiteText, btnSize))
        {
            this.chatPinned = false;
            this.chatWindowOpen = true;
            this.popupChatScrollToBottomPending = true;
        }
        if (ImGui.IsItemHovered())
        {
            ImGui.SetTooltip("Pop up chat");
        }

        var settings = this.GetSheetSettings();
        var sidebarChatMessages = settings.ChatMessages.OrderBy(x => x.TimestampUtc).TakeLast(100).ToList();
        var sidebarSheetId = this.currentSheet.Id ?? string.Empty;
        if (!string.Equals(this.lastSidebarChatSheetId, sidebarSheetId, StringComparison.OrdinalIgnoreCase))
        {
            this.lastSidebarChatSheetId = sidebarSheetId;
            this.sidebarChatScrollToBottomPending = true;
        }
        var height = MathF.Max(240f, ImGui.GetContentRegionAvail().Y - 12f);
        ImGui.PushStyleColor(ImGuiCol.ChildBg, this.GetChatPanelBackgroundColor());
        if (ImGui.BeginChild("##sidebar-chat-scroll", new Vector2(-1f, height - 34f), true))
        {
            foreach (var message in sidebarChatMessages)
            {
                this.DrawChatMessageEntry(message, settings);
            }

            ImGui.Dummy(new Vector2(1f, 1f));
            if (this.sidebarChatScrollToBottomPending || ImGui.IsWindowAppearing())
            {
                ImGui.SetScrollHereY(1f);
                this.sidebarChatScrollToBottomPending = false;
            }
            if (ImGui.IsWindowHovered() && ImGui.IsMouseClicked(ImGuiMouseButton.Left))
            {
                this.ClearUnreadChatAttention();
            }

            var min = ImGui.GetWindowPos();
            var max = min + ImGui.GetWindowSize();
            var drawList = ImGui.GetWindowDrawList();
            drawList.AddRectFilledMultiColor(
                new Vector2(min.X, max.Y - 28f),
                max,
                ImGui.ColorConvertFloat4ToU32(new Vector4(0f, 0f, 0f, 0f)),
                ImGui.ColorConvertFloat4ToU32(new Vector4(0f, 0f, 0f, 0f)),
                ImGui.ColorConvertFloat4ToU32(new Vector4(0.08f, 0.08f, 0.08f, 0.78f)),
                ImGui.ColorConvertFloat4ToU32(new Vector4(0.08f, 0.08f, 0.08f, 0.78f)));
        }
        ImGui.EndChild();
        ImGui.PopStyleColor();

        var sidebarSubmitted = this.DrawPromptedChatInputSingleLine("##sidebar-chat-input-frame", "##sidebar-chat-input", ref this.chatInput, 512);
        this.DrawChatMentionSuggestions("sidebar", ref this.chatInput);
        if (sidebarSubmitted)
        {
            this.TrySendChatMessage(settings);
        }
    }

    private void TrySendChatMessage(ExpandedSheetSettings settings)
    {
        if (this.currentSheet is null || string.IsNullOrWhiteSpace(this.chatInput))
        {
            return;
        }

        var messageText = this.chatInput.Trim();
        this.chatInput = string.Empty;
        var sheetId = this.currentSheet.Id;
        var characterName = this.GetSafeCharacterFullName();
        this.sidebarChatScrollToBottomPending = true;
        this.popupChatScrollToBottomPending = true;
        var localMessage = new SheetChatMessage
        {
            Id = $"local-{Guid.NewGuid():N}",
            AuthorUserId = this.configuration.UserId,
            AuthorName = characterName,
            Message = messageText,
            TimestampUtc = DateTimeOffset.UtcNow,
        };
        settings.ChatMessages.Add(localMessage);
        if (settings.ChatMessages.Count > 200)
        {
            settings.ChatMessages = settings.ChatMessages.OrderBy(x => x.TimestampUtc).TakeLast(200).ToList();
        }

        _ = Task.Run(async () =>
        {
            try
            {
                using var sendCts = CancellationTokenSource.CreateLinkedTokenSource(CancellationToken.None);
                sendCts.CancelAfter(TimeSpan.FromSeconds(15));
                await this.supabase.PostChatMessageAsync(sheetId, characterName, messageText, sendCts.Token).ConfigureAwait(false);
                this.RecordActivity("Chat message sent", null, null, messageText);
                this.lastLiveRefreshUtc = DateTimeOffset.MinValue;
            }
            catch (Exception ex)
            {
                settings.ChatMessages.RemoveAll(x => string.Equals(x.Id, localMessage.Id, StringComparison.OrdinalIgnoreCase));
                this.statusMessage = string.IsNullOrWhiteSpace(ex.Message) ? "Could not send chat message." : ex.Message;
            }
        });
    }


    private bool DrawPromptedChatInputSingleLine(string frameId, string inputId, ref string value, int maxLength)
    {
        var submitted = false;
        ImGui.PushStyleColor(ImGuiCol.ChildBg, new Vector4(0.08f, 0.08f, 0.08f, 0.26f));
        if (ImGui.BeginChild(frameId, new Vector2(-1f, 30f), true, ImGuiWindowFlags.NoScrollbar | ImGuiWindowFlags.NoScrollWithMouse))
        {
            ImGui.SetCursorPos(new Vector2(7f, 5f));
            ImGui.TextColored(new Vector4(1f, 1f, 1f, 0.78f), ">");
            ImGui.SameLine(0f, 5f);
            ImGui.SetNextItemWidth(-1f);
            ImGui.PushStyleColor(ImGuiCol.FrameBg, new Vector4(0f, 0f, 0f, 0f));
            ImGui.PushStyleColor(ImGuiCol.Border, new Vector4(0f, 0f, 0f, 0f));
            submitted = ImGui.InputText(inputId, ref value, maxLength, ImGuiInputTextFlags.EnterReturnsTrue);
            this.chatInputActive = ImGui.IsItemActive();
            if (this.chatInputActive)
            {
                this.ClearUnreadChatAttention();
            }
            ImGui.PopStyleColor(2);
            if (ImGui.IsItemHovered())
            {
                ImGui.SetTooltip("Type a message and press Enter.");
            }
        }
        ImGui.EndChild();
        this.mentionSuggestionAnchorMin = ImGui.GetItemRectMin();
        this.mentionSuggestionAnchorMax = ImGui.GetItemRectMax();
        ImGui.PopStyleColor();
        return submitted;
    }

    private void DrawPromptedChatInputMultiline(string frameId, string inputId, ref string value, int maxLength, Vector2 size)
    {
        ImGui.PushStyleColor(ImGuiCol.ChildBg, new Vector4(0.08f, 0.08f, 0.08f, 0.26f));
        if (ImGui.BeginChild(frameId, size, true))
        {
            ImGui.SetCursorPos(new Vector2(7f, 6f));
            ImGui.TextColored(new Vector4(1f, 1f, 1f, 0.78f), ">");
            ImGui.SetCursorPos(new Vector2(20f, 3f));
            ImGui.PushStyleColor(ImGuiCol.FrameBg, new Vector4(0f, 0f, 0f, 0f));
            ImGui.PushStyleColor(ImGuiCol.Border, new Vector4(0f, 0f, 0f, 0f));
            ImGui.InputTextMultiline(inputId, ref value, maxLength, new Vector2(-1f, size.Y - 8f));
            this.chatInputActive = ImGui.IsItemActive();
            if (this.chatInputActive)
            {
                this.ClearUnreadChatAttention();
            }
            ImGui.PopStyleColor(2);
        }
        ImGui.EndChild();
        this.mentionSuggestionAnchorMin = ImGui.GetItemRectMin();
        this.mentionSuggestionAnchorMax = ImGui.GetItemRectMax();
        ImGui.PopStyleColor();
    }

    private Vector4 GetChatPanelBackgroundColor()
    {
        var normal = new Vector4(0.16f, 0.16f, 0.16f, 0.55f);
        if (!this.chatHasUnreadAttention)
        {
            return normal;
        }

        var blinkOn = ((int)(DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() / 2000L)) % 2 == 0;
        return blinkOn ? new Vector4(0.24f, 0.24f, 0.24f, 0.72f) : normal;
    }

    private void ClearUnreadChatAttention()
    {
        this.unreadChatCount = 0;
        this.chatHasUnreadAttention = false;
    }

    private void DrawChatMentionSuggestions(string popupId, ref string value)
    {
        if (!this.TryGetPendingMention(value, out _, out var query))
        {
            return;
        }

        var suggestions = this.GetMentionSuggestions(query);
        if (suggestions.Count == 0)
        {
            return;
        }

        var width = MathF.Max(220f, this.mentionSuggestionAnchorMax.X - this.mentionSuggestionAnchorMin.X);
        var visibleRows = Math.Min(12, suggestions.Count);
        var desiredHeight = 10f + (visibleRows * 22f);
        var position = new Vector2(this.mentionSuggestionAnchorMin.X, this.mentionSuggestionAnchorMin.Y - desiredHeight - 6f);

        ImGui.SetNextWindowPos(position, ImGuiCond.Always);
        ImGui.SetNextWindowSize(new Vector2(width, desiredHeight), ImGuiCond.Always);
        ImGui.PushStyleVar(ImGuiStyleVar.WindowPadding, new Vector2(6f, 6f));
        ImGui.PushStyleVar(ImGuiStyleVar.WindowRounding, 4f);
        ImGui.PushStyleVar(ImGuiStyleVar.WindowBorderSize, 1f);
        ImGui.PushStyleColor(ImGuiCol.WindowBg, new Vector4(0.10f, 0.10f, 0.10f, 0.98f));
        ImGui.PushStyleColor(ImGuiCol.Border, new Vector4(0.28f, 0.28f, 0.28f, 0.95f));
        var flags = ImGuiWindowFlags.NoDecoration | ImGuiWindowFlags.NoMove | ImGuiWindowFlags.NoSavedSettings | ImGuiWindowFlags.NoFocusOnAppearing | ImGuiWindowFlags.NoNav;
        if (ImGui.Begin($"##mention-suggestions-{popupId}", flags))
        {
            if (ImGui.BeginChild($"##mention-suggestions-scroll-{popupId}", new Vector2(-1f, -1f), false, ImGuiWindowFlags.AlwaysVerticalScrollbar))
            {
                foreach (var suggestion in suggestions)
                {
                    if (ImGui.Selectable(suggestion, false))
                    {
                        this.ApplyMentionSuggestion(ref value, suggestion);
                        this.chatInputActive = true;
                        this.ClearUnreadChatAttention();
                    }
                }
            }
            ImGui.EndChild();
        }
        ImGui.End();
        ImGui.PopStyleColor(2);
        ImGui.PopStyleVar(3);
    }

    private bool TryGetPendingMention(string? input, out int atIndex, out string query)
    {
        atIndex = -1;
        query = string.Empty;
        if (string.IsNullOrWhiteSpace(input))
        {
            return false;
        }

        var lastAt = input.LastIndexOf('@');
        if (lastAt < 0)
        {
            return false;
        }

        if (lastAt > 0 && !char.IsWhiteSpace(input[lastAt - 1]))
        {
            return false;
        }

        var tail = input[(lastAt + 1)..];
        if (tail.Contains(' ') || tail.Contains('\n') || tail.Contains('\r') || tail.Contains('\t'))
        {
            return false;
        }

        atIndex = lastAt;
        query = tail.Trim();
        return true;
    }

    private List<string> GetMentionSuggestions(string query)
    {
        var results = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var settings = this.GetSheetSettings();

        foreach (var profile in settings.MemberProfiles ?? new List<SheetMemberProfile>())
        {
            if (!string.IsNullOrWhiteSpace(profile.CharacterName))
            {
                results.Add(profile.CharacterName.Trim());
            }
        }

        lock (this.currentSheetMembers)
        {
            foreach (var member in this.currentSheetMembers)
            {
                var name = !string.IsNullOrWhiteSpace(member.CharacterName) ? member.CharacterName : member.UserId;
                if (!string.IsNullOrWhiteSpace(name))
                {
                    results.Add(name.Trim());
                }
            }
        }

        foreach (var presence in settings.Presence ?? new List<SheetPresenceInfo>())
        {
            if (!string.IsNullOrWhiteSpace(presence.UserName))
            {
                results.Add(presence.UserName.Trim());
            }
        }

        if (string.IsNullOrWhiteSpace(query))
        {
            return results.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).Take(12).ToList();
        }

        return results
            .Where(name => name.StartsWith(query, StringComparison.OrdinalIgnoreCase) || name.Contains(query, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(name => name.StartsWith(query, StringComparison.OrdinalIgnoreCase))
            .ThenBy(name => name, StringComparer.OrdinalIgnoreCase)
            .Take(12)
            .ToList();
    }

    private void ApplyMentionSuggestion(ref string value, string suggestion)
    {
        if (!this.TryGetPendingMention(value, out var atIndex, out _))
        {
            return;
        }

        value = string.Concat(value.AsSpan(0, atIndex), $"@{suggestion} ");
    }

    private void DrawChatMessageEntry(SheetChatMessage message, ExpandedSheetSettings settings)
    {
        var isOwnerMessage = this.currentSheet is not null && string.Equals(message.AuthorUserId, this.currentSheet.OwnerId, StringComparison.OrdinalIgnoreCase);
        var profile = string.IsNullOrWhiteSpace(message.AuthorUserId) ? null : this.FindMemberProfile(settings, message.AuthorUserId);
        var mention = this.IsMentionedInMessage(message);
        var color = isOwnerMessage
            ? HexToVector4("#d99a32")
            : profile is not null && profile.RoleColor != 0
                ? ImGui.ColorConvertU32ToFloat4(profile.RoleColor)
                : WhiteText;
        if (mention)
        {
            var mentionMin = ImGui.GetCursorScreenPos();
            var mentionMax = mentionMin + new Vector2(ImGui.GetContentRegionAvail().X, 36f);
            ImGui.GetWindowDrawList().AddRectFilled(mentionMin, mentionMax, ImGui.ColorConvertFloat4ToU32(new Vector4(0.42f, 0.34f, 0.08f, 0.18f)), 4f);
        }
        var label = isOwnerMessage ? $"★ {message.AuthorName}" : message.AuthorName;
        var canOpenPermissions = this.CanManageRoles() && !string.IsNullOrWhiteSpace(message.AuthorUserId);

        ImGui.TextColored(color, label);
        if (canOpenPermissions && ImGui.IsItemHovered())
        {
            var rectMin = ImGui.GetItemRectMin();
            var rectMax = ImGui.GetItemRectMax();
            ImGui.GetWindowDrawList().AddLine(new Vector2(rectMin.X, rectMax.Y - 1f), new Vector2(rectMax.X, rectMax.Y - 1f), ImGui.ColorConvertFloat4ToU32(color), 1.2f);
            ImGui.SetMouseCursor(ImGuiMouseCursor.Hand);
            ImGui.SetTooltip($"Open permissions for {message.AuthorName}");
        }
        if (canOpenPermissions && ImGui.IsItemClicked())
        {
            this.OpenPermissionsEditorForUser(message.AuthorUserId);
        }

        ImGui.SameLine(0, 6f);
        ImGui.TextDisabled($"・ {message.TimestampUtc.LocalDateTime:MMM. dd - HH:mm}");
        if (mention)
        {
            ImGui.PushStyleColor(ImGuiCol.Text, HexToVector4("#ffd463"));
            ImGui.TextWrapped(message.Message);
            ImGui.PopStyleColor();
        }
        else
        {
            ImGui.TextWrapped(message.Message);
        }
        ImGui.Spacing();
    }

    private void DrawPresenceInlinePanel(float width, float height)
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var settings = this.GetSheetSettings();
        ImGui.PushStyleColor(ImGuiCol.ChildBg, new Vector4(0.18f, 0.18f, 0.18f, 0.42f));
        if (ImGui.BeginChild("##presence-inline", new Vector2(width, height), true, ImGuiWindowFlags.AlwaysVerticalScrollbar))
        {
            ImGui.TextDisabled("Users seen in this sheet");
            ImGui.Separator();
            this.DrawPresenceEntries(settings);
            var min = ImGui.GetWindowPos();
            var max = min + ImGui.GetWindowSize();
            var drawList = ImGui.GetWindowDrawList();
            drawList.AddRectFilledMultiColor(
                new Vector2(min.X, max.Y - 24f),
                max,
                ImGui.ColorConvertFloat4ToU32(new Vector4(0f, 0f, 0f, 0f)),
                ImGui.ColorConvertFloat4ToU32(new Vector4(0f, 0f, 0f, 0f)),
                ImGui.ColorConvertFloat4ToU32(new Vector4(0.08f, 0.08f, 0.08f, 0.65f)),
                ImGui.ColorConvertFloat4ToU32(new Vector4(0.08f, 0.08f, 0.08f, 0.65f)));
        }
        ImGui.EndChild();
        ImGui.PopStyleColor();
    }

    private Vector4 GetUserAccentColor(string? userId)
    {
        var seed = 23;
        foreach (var ch in userId ?? string.Empty)
        {
            seed = (seed * 31) + ch;
        }

        var palette = new[]
        {
            "#ff9100", "#00b0ff", "#83db18", "#d96cff", "#ff6b6b", "#4dd0e1", "#f4d35e", "#ff8fab",
        };
        return HexToVector4(palette[Math.Abs(seed) % palette.Length]);
    }

    private string GetPresenceStatusLabel(SheetPresenceInfo presence)
    {
        var idleSeconds = Math.Max(0d, (DateTimeOffset.UtcNow - presence.LastSeenUtc).TotalSeconds);
        if (idleSeconds <= 5d)
        {
            return string.IsNullOrWhiteSpace(presence.EditingCellKey) ? "Online" : "Editing";
        }

        if (idleSeconds <= 25d)
        {
            return "Active";
        }

        return $"Idle {Math.Round(idleSeconds)}s";
    }

    private void BuildSearchResults()
    {
        this.searchResults.Clear();
        if (this.currentSheet is null)
        {
            return;
        }

        var document = this.currentSheet.Data;
        document.Normalize(this.currentSheet.RowsCount, this.currentSheet.ColsCount);
        var query = this.searchQuery.Trim();
        if (string.IsNullOrWhiteSpace(query))
        {
            return;
        }

        for (var tabIndex = 0; tabIndex < document.Tabs.Count; tabIndex++)
        {
            var tab = document.Tabs[tabIndex];
            foreach (var pair in tab.Cells)
            {
                if (!pair.Value.Value.Contains(query, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                this.searchResults.Add(new SearchResultItem(tabIndex, tab.Name, pair.Key, ToA1(pair.Key), pair.Value.Value));
            }
        }
    }

    private void ReplaceAllMatches()
    {
        if (this.currentSheet is null || string.IsNullOrWhiteSpace(this.searchQuery))
        {
            return;
        }

        var document = this.currentSheet.Data;
        document.Normalize(this.currentSheet.RowsCount, this.currentSheet.ColsCount);
        var tabs = this.replaceAllTabs ? document.Tabs : new List<SheetTabData> { this.GetActiveTab()! };
        var replacements = 0;
        this.PushUndoSnapshot();
        foreach (var tab in tabs.Where(x => x is not null))
        {
            foreach (var pair in tab.Cells)
            {
                var cell = pair.Value;
                if (cell is null)
                {
                    continue;
                }

                var current = cell.Value ?? string.Empty;
                if (!current.Contains(this.searchQuery, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                cell.Value = ReplaceInvariant(current, this.searchQuery, this.replaceWithQuery);
                replacements++;
            }
        }

        if (replacements > 0)
        {
            this.RecordActivity("Replace all", null, this.searchQuery, this.replaceWithQuery);
            this.statusMessage = $"Replaced {replacements} cell value(s).";
            this.BuildSearchResults();
        }
        else
        {
            this.statusMessage = "No matching cells were found to replace.";
        }
    }

    private static string ReplaceInvariant(string source, string search, string replaceWith)
    {
        if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(search))
        {
            return source;
        }

        return System.Text.RegularExpressions.Regex.Replace(source, System.Text.RegularExpressions.Regex.Escape(search), replaceWith ?? string.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    }

    private void MaybeAutoReconnectLastSheet()
    {
        if (!this.supabase.IsAuthenticated || this.currentSheet is not null || string.IsNullOrWhiteSpace(this.configuration.LastOpenedSheetId) || this.actionInProgress)
        {
            return;
        }

        var now = DateTimeOffset.UtcNow;
        if (this.lastAutoReconnectAttemptUtc != DateTimeOffset.MinValue && (now - this.lastAutoReconnectAttemptUtc).TotalSeconds < 6d)
        {
            return;
        }

        this.lastAutoReconnectAttemptUtc = now;
        _ = this.RunActionAsync(async () =>
        {
            if (this.currentSheet is null && !string.IsNullOrWhiteSpace(this.configuration.LastOpenedSheetId))
            {
                await this.LoadSheetAsync(this.configuration.LastOpenedSheetId).ConfigureAwait(false);
            }
        }, "Could not auto-reconnect the previous sheet.");
    }

    private void MaybeEnforceTemporaryAccessExpiry()
    {
        if (this.currentSheet is null || this.currentRole != SheetAccessRole.Owner || this.actionInProgress)
        {
            return;
        }

        var now = DateTimeOffset.UtcNow;
        if (this.lastTemporaryAccessSweepUtc != DateTimeOffset.MinValue && (now - this.lastTemporaryAccessSweepUtc).TotalSeconds < 10d)
        {
            return;
        }

        this.lastTemporaryAccessSweepUtc = now;
        var settings = this.GetSheetSettings();
        var expiredProfiles = settings.MemberProfiles
            .Where(x => x.AccessExpiresAtUtc.HasValue && x.AccessExpiresAtUtc.Value <= now && !string.Equals(x.UserId, this.currentSheet.OwnerId, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (expiredProfiles.Count == 0)
        {
            return;
        }

        _ = this.RunActionAsync(async () =>
        {
            foreach (var profile in expiredProfiles)
            {
                await this.supabase.RemoveSheetMemberAsync(this.currentSheet.Id, profile.UserId, "Temporary access expired.").ConfigureAwait(false);
            }

            await this.RefreshSheetMembersAsync().ConfigureAwait(false);
            this.statusMessage = $"Expired temporary access for {expiredProfiles.Count} user(s).";
        }, "Could not enforce temporary access expiration.");
    }

    private void TryRestoreDetachedWindowLayout(string kind)
    {
        switch (kind)
        {
            case "chat" when this.configuration.HasSavedChatWindowLayout:
                ImGui.SetWindowPos(new Vector2(this.configuration.ChatWindowPosX, this.configuration.ChatWindowPosY), ImGuiCond.FirstUseEver);
                ImGui.SetWindowSize(new Vector2(MathF.Max(320f, this.configuration.ChatWindowWidth), MathF.Max(220f, this.configuration.ChatWindowHeight)), ImGuiCond.FirstUseEver);
                break;
            case "sheetlist" when this.configuration.HasSavedSheetListWindowLayout:
                ImGui.SetWindowPos(new Vector2(this.configuration.SheetListWindowPosX, this.configuration.SheetListWindowPosY), ImGuiCond.FirstUseEver);
                ImGui.SetWindowSize(new Vector2(MathF.Max(300f, this.configuration.SheetListWindowWidth), MathF.Max(220f, this.configuration.SheetListWindowHeight)), ImGuiCond.FirstUseEver);
                break;
            case "presence" when this.configuration.HasSavedPresenceWindowLayout:
                ImGui.SetWindowPos(new Vector2(this.configuration.PresenceWindowPosX, this.configuration.PresenceWindowPosY), ImGuiCond.FirstUseEver);
                ImGui.SetWindowSize(new Vector2(MathF.Max(300f, this.configuration.PresenceWindowWidth), MathF.Max(220f, this.configuration.PresenceWindowHeight)), ImGuiCond.FirstUseEver);
                break;
        }
    }

    private void TryPersistDetachedWindowLayout(string kind)
    {
        var pos = ImGui.GetWindowPos();
        var size = ImGui.GetWindowSize();
        if (size.X < 1f || size.Y < 1f)
        {
            return;
        }

        var shouldSave = false;
        switch (kind)
        {
            case "chat":
            {
                var posChanged = float.IsNaN(this.lastPersistedChatWindowPos.X) || Vector2.DistanceSquared(pos, this.lastPersistedChatWindowPos) > 0.25f;
                var sizeChanged = float.IsNaN(this.lastPersistedChatWindowSize.X) || Vector2.DistanceSquared(size, this.lastPersistedChatWindowSize) > 0.25f;
                if (!posChanged && !sizeChanged)
                {
                    return;
                }

                this.lastPersistedChatWindowPos = pos;
                this.lastPersistedChatWindowSize = size;
                this.configuration.ChatWindowPosX = pos.X;
                this.configuration.ChatWindowPosY = pos.Y;
                this.configuration.ChatWindowWidth = size.X;
                this.configuration.ChatWindowHeight = size.Y;
                this.configuration.HasSavedChatWindowLayout = true;
                shouldSave = true;
                break;
            }
            case "sheetlist":
            {
                var posChanged = float.IsNaN(this.lastPersistedSheetListWindowPos.X) || Vector2.DistanceSquared(pos, this.lastPersistedSheetListWindowPos) > 0.25f;
                var sizeChanged = float.IsNaN(this.lastPersistedSheetListWindowSize.X) || Vector2.DistanceSquared(size, this.lastPersistedSheetListWindowSize) > 0.25f;
                if (!posChanged && !sizeChanged)
                {
                    return;
                }

                this.lastPersistedSheetListWindowPos = pos;
                this.lastPersistedSheetListWindowSize = size;
                this.configuration.SheetListWindowPosX = pos.X;
                this.configuration.SheetListWindowPosY = pos.Y;
                this.configuration.SheetListWindowWidth = size.X;
                this.configuration.SheetListWindowHeight = size.Y;
                this.configuration.HasSavedSheetListWindowLayout = true;
                shouldSave = true;
                break;
            }
            case "presence":
            {
                var posChanged = float.IsNaN(this.lastPersistedPresenceWindowPos.X) || Vector2.DistanceSquared(pos, this.lastPersistedPresenceWindowPos) > 0.25f;
                var sizeChanged = float.IsNaN(this.lastPersistedPresenceWindowSize.X) || Vector2.DistanceSquared(size, this.lastPersistedPresenceWindowSize) > 0.25f;
                if (!posChanged && !sizeChanged)
                {
                    return;
                }

                this.lastPersistedPresenceWindowPos = pos;
                this.lastPersistedPresenceWindowSize = size;
                this.configuration.PresenceWindowPosX = pos.X;
                this.configuration.PresenceWindowPosY = pos.Y;
                this.configuration.PresenceWindowWidth = size.X;
                this.configuration.PresenceWindowHeight = size.Y;
                this.configuration.HasSavedPresenceWindowLayout = true;
                shouldSave = true;
                break;
            }
        }

        if (shouldSave)
        {
            this.configuration.Save();
        }
    }

    private void MaybeRefreshSheetList()
    {
        if (!this.supabase.IsAuthenticated || this.actionInProgress)
        {
            return;
        }

        var now = DateTimeOffset.UtcNow;
        if (this.lastSheetListRefreshUtc != DateTimeOffset.MinValue && (now - this.lastSheetListRefreshUtc).TotalSeconds < 30.0d)
        {
            return;
        }

        this.lastSheetListRefreshUtc = now;
        _ = Task.Run(async () =>
        {
            try
            {
                await this.ReloadSheetListAsync().ConfigureAwait(false);
            }
            catch
            {
            }
        });
    }

    private void MaybeAutoSave()
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var settings = this.GetSheetSettings();
        settings.AutoSaveLocal = false;
        settings.AutoSaveCloud = true;
        settings.AutoSaveSeconds = 20;

        var now = DateTimeOffset.UtcNow;
        if (this.lastAutoSaveTickUtc != DateTimeOffset.MinValue && (now - this.lastAutoSaveTickUtc).TotalSeconds < settings.AutoSaveSeconds)
        {
            return;
        }

        if (this.autoSaveBackoffUntilUtc > now)
        {
            return;
        }

        this.lastAutoSaveTickUtc = now;
        if (!settings.AutoSaveCloud || !this.supabase.IsAuthenticated || !this.CanEditCurrentSheet || this.actionInProgress || this.liveRefreshInFlight)
        {
            return;
        }

        if (!this.HasPersistableChangesSinceLastSave())
        {
            return;
        }

        if (this.lastPersistableChangeUtc != DateTimeOffset.MinValue && (now - this.lastPersistableChangeUtc).TotalMilliseconds < 1800d)
        {
            return;
        }

        if (this.autoSaveInFlight)
        {
            return;
        }

        this.suppressNextCloudSaveHistory = true;
        this.suppressNextCloudSaveStatus = true;
        this.autoSaveInFlight = true;
        _ = Task.Run(async () =>
        {
            try
            {
                using var autoSaveCts = CancellationTokenSource.CreateLinkedTokenSource(CancellationToken.None);
                autoSaveCts.CancelAfter(TimeSpan.FromSeconds(75));
                await this.SaveCurrentSheetToCloudAsync(autoSaveCts.Token).ConfigureAwait(false);
                this.autoSaveBackoffUntilUtc = DateTimeOffset.MinValue;
            }
            catch (Exception ex)
            {
                var message = string.IsNullOrWhiteSpace(ex.Message) ? "Cloud autosave failed." : ex.Message;
                this.statusMessage = message;
                this.autoSaveBackoffUntilUtc = message.Contains("expired", StringComparison.OrdinalIgnoreCase)
                    || message.Contains("Unauthorized", StringComparison.OrdinalIgnoreCase)
                    || message.Contains("401", StringComparison.OrdinalIgnoreCase)
                    ? DateTimeOffset.UtcNow.AddMinutes(3)
                    : DateTimeOffset.UtcNow.AddSeconds(75);
                Plugin.Log.Warning(ex, "Cloud autosave failed.");
            }
            finally
            {
                this.autoSaveInFlight = false;
            }
        });
    }


    private void MaybeSyncPresenceHeartbeat()
    {
        if (this.currentSheet is null || !this.supabase.IsAuthenticated)
        {
            return;
        }

        var now = DateTimeOffset.UtcNow;
        if (this.autoSaveInFlight || (this.lastPresenceHeartbeatUtc != DateTimeOffset.MinValue && (now - this.lastPresenceHeartbeatUtc).TotalSeconds < 5.0d))
        {
            return;
        }

        this.lastPresenceHeartbeatUtc = now;
        var sheetId = this.currentSheet.Id;
        var characterName = this.GetSafeCharacterFullName();
        var activeTabName = this.GetActiveTab()?.Name;
        var editingKey = this.editingCellKey ?? this.selectedCellKey;
        _ = Task.Run(async () =>
        {
            try
            {
                await this.supabase.SyncSheetPresenceAsync(sheetId, characterName, activeTabName, editingKey).ConfigureAwait(false);
            }
            catch (SupabaseApiException ex) when ((int)ex.StatusCode == 403 || (int)ex.StatusCode == 404)
            {
                this.removedByAdminMessage = "You were removed by a admin.";
                this.HandleLiveSheetRemoval(sheetId);
            }
            catch
            {
            }
        });
    }

    private void MaybeRefreshLiveSheetState()
    {
        if (this.currentSheet is null || !this.supabase.IsAuthenticated || this.liveRefreshInFlight || this.autoSaveInFlight)
        {
            return;
        }

        var now = DateTimeOffset.UtcNow;
        if (this.lastLiveRefreshUtc != DateTimeOffset.MinValue && (now - this.lastLiveRefreshUtc).TotalSeconds < 2.0d)
        {
            return;
        }

        this.lastLiveRefreshUtc = now;
        this.liveRefreshInFlight = true;
        var sheetId = this.currentSheet.Id;
        _ = Task.Run(async () =>
        {
            try
            {
                await this.RefreshLiveSheetStateAsync(sheetId).ConfigureAwait(false);
            }
            finally
            {
                this.liveRefreshInFlight = false;
            }
        });
    }

    private async Task RefreshLiveSheetStateAsync(string sheetId)
    {
        try
        {
            var runtime = await this.supabase.GetSheetRuntimeStateAsync(sheetId).ConfigureAwait(false);
            this.MergeRuntimeState(runtime);

            if (this.currentSheet is null)
            {
                return;
            }

            var needsFullRefresh = this.currentSheet.Version != runtime.Version;
            if (!needsFullRefresh)
            {
                return;
            }

            var remote = await this.supabase.GetSheetAsync(sheetId).ConfigureAwait(false);
            if (remote is null)
            {
                this.HandleLiveSheetRemoval(sheetId);
                return;
            }

            remote.Data.Normalize(remote.RowsCount, remote.ColsCount);
            var remoteRole = await this.supabase.GetAccessRoleAsync(sheetId).ConfigureAwait(false);
            this.MergeLiveSheetState(remote, remoteRole);
            this.MergeRuntimeState(runtime);
        }
        catch (SupabaseApiException ex) when ((int)ex.StatusCode == 403 || (int)ex.StatusCode == 404)
        {
            this.HandleLiveSheetRemoval(sheetId);
        }
        catch
        {
        }
    }

    private void MergeRuntimeState(SheetRuntimeState runtime)
    {
        if (this.currentSheet is null || runtime is null || !string.Equals(this.currentSheet.Id, runtime.SheetId, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var settings = this.GetSheetSettings();
        settings.Presence = runtime.Presence
            .OrderByDescending(x => x.LastSeenUtc)
            .Take(64)
            .ToList();
        var mergedChatMessages = this.MergeRuntimeChatMessages(settings.ChatMessages, runtime.ChatMessages)
            .OrderBy(x => x.TimestampUtc)
            .TakeLast(200)
            .ToList();
        if (!ReferenceEquals(settings.ChatMessages, mergedChatMessages))
        {
            settings.ChatMessages = mergedChatMessages;
        }

        this.HandleChatNotification(mergedChatMessages);

        List<SheetMemberRow> runtimeMemberSnapshot;
        lock (this.currentSheetMembers)
        {
            if (runtime.Members is { Count: > 0 })
            {
                this.currentSheetMembers.Clear();
                this.currentSheetMembers.AddRange(runtime.Members.OrderBy(x => x.CreatedAt).ThenBy(x => x.UserId, StringComparer.OrdinalIgnoreCase));
                this.lastSheetMembersRefreshUtc = DateTimeOffset.UtcNow;
            }
            runtimeMemberSnapshot = this.currentSheetMembers.ToList();
        }

        lock (this.currentCellLocks)
        {
            this.currentCellLocks.Clear();
            this.currentCellLocks.AddRange(runtime.CellLocks
                .Where(x => x.ExpiresAt > DateTimeOffset.UtcNow)
                .OrderBy(x => x.CellKey, StringComparer.OrdinalIgnoreCase));
        }

        if (this.CanUseInviteCodes())
        {
            this.currentOwnerUniqueCode = runtime.CurrentUniqueCode ?? this.currentOwnerUniqueCode;
        }
        else
        {
            this.currentOwnerUniqueCode = null;
        }

        foreach (var row in runtimeMemberSnapshot)
        {
            var fallbackName = !string.IsNullOrWhiteSpace(row.CharacterName) ? row.CharacterName : this.ResolveMemberCharacterName(settings, row.UserId);
            var profile = this.GetOrCreateMemberProfile(settings, row.UserId, fallbackName, row.CreatedAt);
            if (!string.IsNullOrWhiteSpace(row.CharacterName))
            {
                profile.CharacterName = row.CharacterName;
            }
            profile.LastSeenUtc = row.LastSeenUtc ?? profile.LastSeenUtc;
            if (!string.IsNullOrWhiteSpace(row.AssignedPresetName))
            {
                profile.AssignedPresetName = row.AssignedPresetName;
            }
            if (row.RoleColor != 0)
            {
                profile.RoleColor = row.RoleColor;
            }
            profile.IsBlocked = row.IsBlocked;
        }

        var myProfile = this.FindMemberProfile(settings, this.configuration.UserId);
        if (!string.IsNullOrWhiteSpace(runtime.RequesterRole))
        {
            this.currentRole = runtime.RequesterRole.Trim().ToLowerInvariant() switch
            {
                "owner" => SheetAccessRole.Owner,
                "editor" => SheetAccessRole.Editor,
                _ => SheetAccessRole.Viewer,
            };
        }
        if (myProfile is not null)
        {
            if (!string.IsNullOrWhiteSpace(runtime.RequesterRoleName))
            {
                myProfile.AssignedPresetName = runtime.RequesterRoleName;
            }
            if (runtime.RequesterRoleColor != 0)
            {
                myProfile.RoleColor = runtime.RequesterRoleColor;
            }
        }

        if (myProfile?.IsBlocked == true)
        {
            this.removedByAdminMessage = "You were removed by a admin.";
            this.HandleLiveSheetRemoval(runtime.SheetId);
            return;
        }

        if (myProfile?.AccessExpiresAtUtc.HasValue == true && myProfile.AccessExpiresAtUtc.Value <= DateTimeOffset.UtcNow && this.currentRole != SheetAccessRole.Owner)
        {
            this.removedByAdminMessage = $"Your temporary access for the sheet {this.currentSheet.Title}, expired.";
            this.HandleLiveSheetRemoval(runtime.SheetId);
            return;
        }

        var newRoleName = !string.IsNullOrWhiteSpace(runtime.RequesterRoleName)
            ? runtime.RequesterRoleName
            : this.GetDisplayedRoleLabel();
        var newRoleColor = this.currentRole == SheetAccessRole.Owner
            ? 0xFFD99A32u
            : (runtime.RequesterRoleColor != 0 ? runtime.RequesterRoleColor : (myProfile?.RoleColor ?? 0xFF5C5C5Cu));
        var newRoleSignature = this.BuildObservedRoleSignature(newRoleName, newRoleColor);
        if (!this.hasObservedRemoteRoleState)
        {
            this.hasObservedRemoteRoleState = true;
            this.lastObservedRemoteRoleSignature = newRoleSignature;
        }
        else if (this.currentRole != SheetAccessRole.Owner && !string.Equals(this.lastObservedRemoteRoleSignature, newRoleSignature, StringComparison.Ordinal))
        {
            var ownerProfile = this.FindMemberProfile(settings, this.currentSheet.OwnerId);
            var ownerName = ownerProfile?.CharacterName;
            if (string.IsNullOrWhiteSpace(ownerName))
            {
                ownerName = this.ResolveMemberCharacterName(settings, this.currentSheet.OwnerId);
            }

            this.QueueRoleChangedPopup(ownerName ?? "Sheet owner", newRoleName, newRoleColor);
            this.lastObservedRemoteRoleSignature = newRoleSignature;
        }

        lock (this.accessibleSheetsSync)
        {
            var summary = this.accessibleSheets.FirstOrDefault(x => string.Equals(x.Id, runtime.SheetId, StringComparison.OrdinalIgnoreCase));
            if (summary is not null)
            {
                summary.Version = Math.Max(summary.Version, runtime.Version);
                summary.UpdatedAt = runtime.UpdatedAt;
                summary.UserRole = this.GetDisplayedRoleLabel();
                summary.UserRoleColor = this.currentRole == SheetAccessRole.Owner
                    ? 0xFFD99A32u
                    : (myProfile?.RoleColor ?? 0u);
            }
        }
    }

    private List<SheetChatMessage> MergeRuntimeChatMessages(List<SheetChatMessage>? currentMessages, List<SheetChatMessage>? runtimeMessages)
    {
        var combined = new List<SheetChatMessage>();
        combined.AddRange(this.FilterVisibleChatMessages(currentMessages));

        foreach (var runtimeMessage in this.FilterVisibleChatMessages(runtimeMessages))
        {
            if (runtimeMessage is null)
            {
                continue;
            }

            combined.RemoveAll(existing => this.IsTemporaryChatEcho(existing, runtimeMessage));
            combined.Add(runtimeMessage);
        }

        var map = new Dictionary<string, SheetChatMessage>(StringComparer.OrdinalIgnoreCase);
        foreach (var message in combined)
        {
            if (message is null)
            {
                continue;
            }

            var key = string.IsNullOrWhiteSpace(message.Id) || message.Id.StartsWith("local-", StringComparison.OrdinalIgnoreCase)
                ? $"{message.AuthorUserId}|{message.TimestampUtc.UtcDateTime.Ticks}|{message.Message}"
                : message.Id;
            map[key] = message;
        }

        return this.FilterVisibleChatMessages(map.Values)
            .OrderBy(x => x.TimestampUtc)
            .TakeLast(200)
            .ToList();
    }

    private IEnumerable<SheetChatMessage> FilterVisibleChatMessages(IEnumerable<SheetChatMessage>? messages)
    {
        var visibleSinceUtc = this.GetCurrentUserChatVisibleSinceUtc();
        return (messages ?? Enumerable.Empty<SheetChatMessage>())
            .Where(message => message is not null && message.TimestampUtc >= visibleSinceUtc);
    }

    private DateTimeOffset GetCurrentUserChatVisibleSinceUtc()
    {
        var visibleSinceUtc = GetCurrentEstChatResetUtc();

        if (this.currentSheet is null)
        {
            return visibleSinceUtc;
        }

        DateTimeOffset? joinedAtUtc = null;
        if (string.Equals(this.currentSheet.OwnerId, this.configuration.UserId, StringComparison.OrdinalIgnoreCase))
        {
            joinedAtUtc = this.currentSheet.CreatedAt;
        }
        else
        {
            lock (this.currentSheetMembers)
            {
                joinedAtUtc = this.currentSheetMembers.FirstOrDefault(member => string.Equals(member.UserId, this.configuration.UserId, StringComparison.OrdinalIgnoreCase))?.CreatedAt;
            }

            if (!joinedAtUtc.HasValue)
            {
                joinedAtUtc = this.FindMemberProfile(this.currentSheet.Data.Settings, this.configuration.UserId)?.JoinedAtUtc;
            }
        }

        if (joinedAtUtc.HasValue && joinedAtUtc.Value > visibleSinceUtc)
        {
            visibleSinceUtc = joinedAtUtc.Value;
        }

        return visibleSinceUtc;
    }

    private static DateTimeOffset GetCurrentEstChatResetUtc()
    {
        var estOffset = TimeSpan.FromHours(-5);
        var estNow = DateTimeOffset.UtcNow.ToOffset(estOffset);
        return new DateTimeOffset(estNow.Year, estNow.Month, estNow.Day, 0, 0, 0, estOffset).ToUniversalTime();
    }

    private bool IsTemporaryChatEcho(SheetChatMessage existing, SheetChatMessage incoming)
    {
        if (existing is null || incoming is null)
        {
            return false;
        }

        var existingIsLocal = !string.IsNullOrWhiteSpace(existing.Id) && existing.Id.StartsWith("local-", StringComparison.OrdinalIgnoreCase);
        var incomingIsLocal = !string.IsNullOrWhiteSpace(incoming.Id) && incoming.Id.StartsWith("local-", StringComparison.OrdinalIgnoreCase);
        if (!existingIsLocal && !incomingIsLocal)
        {
            return false;
        }

        return string.Equals(existing.AuthorUserId, incoming.AuthorUserId, StringComparison.OrdinalIgnoreCase)
            && string.Equals(existing.Message, incoming.Message, StringComparison.Ordinal)
            && Math.Abs((existing.TimestampUtc - incoming.TimestampUtc).TotalSeconds) <= 10d;
    }

    private void MergeLiveSheetState(RemoteSheet remote, SheetAccessRole role)
    {
        if (this.currentSheet is null || !string.Equals(this.currentSheet.Id, remote.Id, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        this.EnsureAdvancedDefaults(this.currentSheet.Data);
        SheetSerializationHelper.SanitizeLoadedDocumentInPlace(remote.Data);
        this.EnsureAdvancedDefaults(remote.Data);

        var previousSettings = this.currentSheet.Data.Settings;
        var preservedTabName = this.GetActiveTab()?.Name;
        var tabWasRemoved = !string.IsNullOrWhiteSpace(preservedTabName)
            && !remote.Data.Tabs.Any(tab => string.Equals(tab.Name, preservedTabName, StringComparison.OrdinalIgnoreCase));

        this.currentSheet.Title = remote.Title;
        this.currentSheet.Code = remote.Code;
        this.currentSheet.RowsCount = remote.RowsCount;
        this.currentSheet.ColsCount = remote.ColsCount;
        this.currentSheet.DefaultRole = remote.DefaultRole;
        this.currentSheet.Version = remote.Version;
        this.currentSheet.CreatedAt = remote.CreatedAt;
        this.currentSheet.UpdatedAt = remote.UpdatedAt;
        this.currentSheet.Data = remote.Data;
        this.EnsureAdvancedDefaults(this.currentSheet.Data);

        if (previousSettings is not null)
        {
            this.currentSheet.Data.Settings.Presence = previousSettings.Presence ?? new List<SheetPresenceInfo>();
            this.currentSheet.Data.Settings.ChatMessages = this.MergeRuntimeChatMessages(previousSettings.ChatMessages, this.currentSheet.Data.Settings.ChatMessages)
                .OrderBy(x => x.TimestampUtc)
                .TakeLast(200)
                .ToList();
        }

        if (!string.IsNullOrWhiteSpace(preservedTabName))
        {
            var restoredIndex = this.currentSheet.Data.Tabs.FindIndex(tab => string.Equals(tab.Name, preservedTabName, StringComparison.OrdinalIgnoreCase));
            if (restoredIndex >= 0)
            {
                this.currentSheet.Data.ActiveTabIndex = restoredIndex;
            }
        }

        if (this.currentSheet.Data.Tabs.Count > 0)
        {
            this.currentSheet.Data.ActiveTabIndex = Math.Clamp(this.currentSheet.Data.ActiveTabIndex, 0, this.currentSheet.Data.Tabs.Count - 1);
        }

        if (tabWasRemoved)
        {
            this.ReleaseCellLockIfNeeded(this.editingCellKey);
            this.editingCellKey = null;
            this.editingCellBuffer = string.Empty;
            this.editingFocusRequested = false;
            this.ClearSelection();
            this.SelectSingleCell("R1C1");
            this.statusMessage = "Your current tab was removed. Switched to another tab.";
        }

        this.EnsureSelectionInBounds();
        this.currentRole = role;
        this.sidebarChatScrollToBottomPending = true;
        this.popupChatScrollToBottomPending = true;
        this.SyncPersistedSheetFingerprint();
        this.SyncChatNotificationAnchor(this.currentSheet.Data.Settings.ChatMessages);

        var localSettings = this.currentSheet.Data.Settings;
        var myProfile = this.FindMemberProfile(localSettings, this.configuration.UserId);
        if (myProfile?.IsBlocked == true)
        {
            this.HandleLiveSheetRemoval(remote.Id);
        }
    }

    private void HandleLiveSheetRemoval(string sheetId)
    {
        if (this.currentSheet is null || !string.Equals(this.currentSheet.Id, sheetId, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        this.configuration.LastOpenedSheetId = string.Empty;
        this.configuration.Save();
        lock (this.accessibleSheetsSync)
        {
            this.accessibleSheets.RemoveAll(x => string.Equals(x.Id, sheetId, StringComparison.OrdinalIgnoreCase));
        }
        this.currentSheet = null;
        this.currentRole = SheetAccessRole.Viewer;
        this.currentSheetMembers.Clear();
        this.currentCellLocks.Clear();
        this.currentSheetBlocklist.Clear();
        this.currentOwnerUniqueCode = null;
        this.hideSharedCode = false;
        this.pendingTabDeletePopupPositions.Clear();
        this.selectedSheetPresetIndex = 0;
        this.saveSheetPresetName = string.Empty;
        this.permissionsWindowOpen = false;
        this.blocklistWindowOpen = false;
        this.permissionsEditorUserId = string.Empty;
        this.pendingRemovalReasonUserId = string.Empty;
        this.pendingRemovalReasonText = string.Empty;
        this.historyWindowOpen = false;
        this.chatWindowOpen = false;
        this.lastSidebarChatSheetId = string.Empty;
        this.lastPopupChatSheetId = string.Empty;
        this.sidebarChatScrollToBottomPending = true;
        this.popupChatScrollToBottomPending = true;
        this.roleChangedPopupOpen = false;
        this.hasObservedRemoteRoleState = false;
        this.lastObservedRemoteRoleSignature = string.Empty;
        this.lastPersistedSheetFingerprint = string.Empty;
        this.lastObservedChatNotificationAnchor = string.Empty;
        this.presenceWindowOpen = false;
        this.presencePinned = true;
        this.ClearSelection();
        this.removedByAdminPopupOpen = true;
        this.statusMessage = this.removedByAdminMessage;
    }

    private void DrawRemovedByAdminPopup()
    {
        if (!this.removedByAdminPopupOpen)
        {
            return;
        }

        var open = this.removedByAdminPopupOpen;
        ImGui.SetNextWindowSize(new Vector2(320f, 130f), ImGuiCond.Appearing);
        if (ImGui.Begin("Removed from sheet###EZSheetsRemovedPopup", ref open, ImGuiWindowFlags.NoCollapse | ImGuiWindowFlags.NoResize))
        {
            ImGui.TextWrapped(this.removedByAdminMessage);
            ImGui.Spacing();
            if (this.DrawStyledButton("removed-ok", "OK", AccentBlue, WhiteText, new Vector2(72f, 24f)))
            {
                open = false;
            }
        }
        ImGui.End();
        this.removedByAdminPopupOpen = open;
    }



    private string BuildObservedRoleSignature(string roleName, uint roleColor)
        => $"{roleName}|{roleColor}";

    private void QueueRoleChangedPopup(string ownerName, string roleName, uint roleColor)
    {
        this.roleChangedOwnerName = string.IsNullOrWhiteSpace(ownerName) ? "Sheet owner" : ownerName;
        this.roleChangedRoleName = string.IsNullOrWhiteSpace(roleName) ? "Viewer" : roleName;
        this.roleChangedRoleColor = roleColor == 0 ? 0xFF5C5C5Cu : roleColor;
        this.roleChangedPopupOpen = true;
    }

    private void DrawRoleChangedPopup()
    {
        if (!this.roleChangedPopupOpen)
        {
            return;
        }

        var open = this.roleChangedPopupOpen;
        ImGui.SetNextWindowSize(new Vector2(410f, 150f), ImGuiCond.Appearing);
        if (ImGui.Begin("Role changed###EZSheetsRoleChangedPopup", ref open, ImGuiWindowFlags.NoCollapse | ImGuiWindowFlags.NoResize))
        {
            ImGui.TextColored(HexToVector4("#d99a32"), this.roleChangedOwnerName);
            ImGui.SameLine(0f, 4f);
            ImGui.TextUnformatted("changed your role to");
            ImGui.SameLine(0f, 4f);
            ImGui.TextColored(ImGui.ColorConvertU32ToFloat4(this.roleChangedRoleColor), this.roleChangedRoleName);
            ImGui.TextUnformatted(".");
            ImGui.Spacing();
            if (this.DrawStyledButton("role-changed-ok", "OK", AccentBlue, WhiteText, new Vector2(72f, 24f)))
            {
                open = false;
            }
        }
        ImGui.End();
        this.roleChangedPopupOpen = open;
    }

    private void DrawBlockedJoinPopup()
    {
        if (!this.blockedJoinPopupOpen)
        {
            return;
        }

        var open = this.blockedJoinPopupOpen;
        ImGui.SetNextWindowSize(new Vector2(320f, 130f), ImGuiCond.Appearing);
        if (ImGui.Begin("Blocked from sheet###EZSheetsBlockedJoinPopup", ref open, ImGuiWindowFlags.NoCollapse | ImGuiWindowFlags.NoResize))
        {
            ImGui.TextWrapped(this.blockedJoinMessage);
            ImGui.Spacing();
            if (this.DrawStyledButton("blocked-join-ok", "OK", AccentBlue, WhiteText, new Vector2(72f, 24f)))
            {
                open = false;
            }
        }
        ImGui.End();
        this.blockedJoinPopupOpen = open;
    }

    private void RecordActivity(string action, string? cellKey, string? oldValue, string? newValue)
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var settings = this.GetSheetSettings();
        settings.ActivityLog.Add(new ActivityEntry
        {
            Action = action,
            CellKey = cellKey,
            OldValue = oldValue,
            NewValue = newValue,
            TabName = this.GetActiveTab()?.Name,
            UserId = this.configuration.UserId,
            UserName = this.GetSafeCharacterFullName(),
            TimestampUtc = DateTimeOffset.UtcNow,
        });

        if (settings.ActivityLog.Count > 300)
        {
            settings.ActivityLog.RemoveRange(0, settings.ActivityLog.Count - 300);
        }
    }

    private void CreateSnapshot()
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var settings = this.GetSheetSettings();
        settings.Snapshots.Add(new SheetVersionSnapshot
        {
            Label = $"Snapshot {DateTime.Now:yyyy-MM-dd HH:mm}",
            CreatedByUserId = this.configuration.UserId,
            CreatedByName = this.GetSafeCharacterFullName(),
            CreatedAtUtc = DateTimeOffset.UtcNow,
            DocumentJson = SheetSerializationHelper.SerializeForSnapshot(this.currentSheet.Data, this.currentSheet.RowsCount, this.currentSheet.ColsCount),
        });
        this.RecordActivity("Snapshot created", null, null, null);
        this.statusMessage = "Snapshot created.";
    }

    private void RestoreSnapshot(string snapshotId)
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var settings = this.GetSheetSettings();
        var snapshot = settings.Snapshots.FirstOrDefault(x => x.Id == snapshotId);
        if (snapshot is null)
        {
            return;
        }

        this.PushUndoSnapshot();
        var restored = JsonSerializer.Deserialize<SheetDocument>(snapshot.DocumentJson) ?? SheetDocument.CreateDefault();
        restored.Normalize(this.currentSheet.RowsCount, this.currentSheet.ColsCount);
        this.currentSheet.Data = restored;
        this.RecordActivity("Snapshot restored", null, null, snapshot.Label);
        this.statusMessage = $"Restored {snapshot.Label}.";
    }

    private void DuplicateActiveTab()
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var document = this.currentSheet.Data;
        var activeTab = this.GetActiveTab();
        if (activeTab is null)
        {
            return;
        }

        this.PushUndoSnapshot();
        var duplicate = JsonSerializer.Deserialize<SheetTabData>(JsonSerializer.Serialize(activeTab)) ?? new SheetTabData();
        duplicate.Name = activeTab.Name + " Copy";
        document.Tabs.Add(duplicate);
        document.ActiveTabIndex = document.Tabs.Count - 1;
        this.RecordActivity("Tab duplicated", null, null, duplicate.Name);
    }

    private void UpdatePresence()
    {
        if (this.currentSheet is null || string.IsNullOrWhiteSpace(this.configuration.UserId))
        {
            return;
        }

        this.MarkPresence(this.currentSheet.Data, "Active");
    }

    private void MarkPresence(SheetDocument document, string reason)
    {
        var settings = document.Settings;
        var entry = settings.Presence.FirstOrDefault(x => x.UserId == this.configuration.UserId);
        if (entry is null)
        {
            entry = new SheetPresenceInfo { UserId = this.configuration.UserId };
            settings.Presence.Add(entry);
        }

        entry.UserName = this.GetSafeCharacterFullName();
        var profile = this.GetOrCreateMemberProfile(settings, this.configuration.UserId, entry.UserName, this.currentSheet?.CreatedAt);
        profile.CharacterName = entry.UserName;
        profile.LastSeenUtc = entry.LastSeenUtc;
        entry.ActiveTabName = this.GetActiveTab()?.Name;
        entry.EditingCellKey = this.editingCellKey ?? this.selectedCellKey;
        entry.LastSeenUtc = DateTimeOffset.UtcNow;
        _ = reason;
    }

    private SheetCellEditLock? FindActiveCellLock(string cellKey)
    {
        lock (this.currentCellLocks)
        {
            return this.currentCellLocks.FirstOrDefault(x => string.Equals(x.CellKey, cellKey, StringComparison.OrdinalIgnoreCase) && x.ExpiresAt > DateTimeOffset.UtcNow);
        }
    }

    private bool TryGetBlockingCellLock(string cellKey, out SheetCellEditLock? cellLock)
    {
        cellLock = this.FindActiveCellLock(cellKey);
        return cellLock is not null && !string.Equals(cellLock.UserId, this.configuration.UserId, StringComparison.OrdinalIgnoreCase);
    }

    private void TryBeginEditingCell(string cellKey, string initialValue)
    {
        if (this.currentSheet is null)
        {
            return;
        }

        if (this.TryGetBlockingCellLock(cellKey, out var blockingLock) && blockingLock is not null)
        {
            this.statusMessage = $"{blockingLock.UserName} is using this cell..";
            return;
        }

        var sheetId = this.currentSheet.Id;
        var characterName = this.GetSafeCharacterFullName();
        this.BeginEditingCell(cellKey, initialValue);
        _ = Task.Run(async () =>
        {
            try
            {
                await this.supabase.AcquireCellLockAsync(sheetId, cellKey, characterName).ConfigureAwait(false);
            }
            catch (SupabaseApiException ex)
            {
                if (string.Equals(this.editingCellKey, cellKey, StringComparison.OrdinalIgnoreCase))
                {
                    this.editingCellKey = null;
                    this.editingCellBuffer = string.Empty;
                    this.editingFocusRequested = false;
                    this.pendingMergedEditOverlay = null;
                }
                this.statusMessage = string.IsNullOrWhiteSpace(ex.Message) ? "That cell is currently in use." : ex.Message;
            }
            catch
            {
                if (string.Equals(this.editingCellKey, cellKey, StringComparison.OrdinalIgnoreCase))
                {
                    this.editingCellKey = null;
                    this.editingCellBuffer = string.Empty;
                    this.editingFocusRequested = false;
                    this.pendingMergedEditOverlay = null;
                }
                this.statusMessage = "That cell is currently in use.";
            }
        });
    }

    private void ReleaseCellLockIfNeeded(string? cellKey)
    {
        if (this.currentSheet is null || string.IsNullOrWhiteSpace(cellKey))
        {
            return;
        }

        var sheetId = this.currentSheet.Id;
        var releaseKey = cellKey!;
        lock (this.currentCellLocks)
        {
            this.currentCellLocks.RemoveAll(x => string.Equals(x.CellKey, releaseKey, StringComparison.OrdinalIgnoreCase) && string.Equals(x.UserId, this.configuration.UserId, StringComparison.OrdinalIgnoreCase));
        }

        _ = Task.Run(async () =>
        {
            try
            {
                await this.supabase.ReleaseCellLockAsync(sheetId, releaseKey).ConfigureAwait(false);
            }
            catch
            {
            }
        });
    }

    private bool CanModifyCell(string key, bool allowLockedCellOverride = false)
    {
        if (!this.CanEditCurrentSheet || this.currentSheet is null)
        {
            return false;
        }

        var activeTab = this.GetActiveTab();
        if (activeTab is null)
        {
            return false;
        }

        var effectiveKey = key;
        if (activeTab.TryGetCell(key, out var mergedCell) && mergedCell is not null && !string.IsNullOrWhiteSpace(mergedCell.MergedInto))
        {
            effectiveKey = mergedCell.MergedInto!;
        }

        if (!allowLockedCellOverride && activeTab.TryGetCell(effectiveKey, out var cell) && cell is not null && cell.Locked)
        {
            return false;
        }

        if (this.currentRole == SheetAccessRole.Owner)
        {
            return true;
        }

        var settings = this.GetSheetSettings();
        if (!TryParseCellKey(key, out var row, out var col))
        {
            return true;
        }

        foreach (var range in settings.ProtectedRanges)
        {
            if (row < Math.Min(range.StartRow, range.EndRow) || row > Math.Max(range.StartRow, range.EndRow)
                || col < Math.Min(range.StartColumn, range.EndColumn) || col > Math.Max(range.StartColumn, range.EndColumn))
            {
                continue;
            }

            if (range.AllowOwnerOnly)
            {
                return false;
            }

            if (!range.EditorUserIds.Contains(this.configuration.UserId))
            {
                return false;
            }
        }

        return true;
    }

    private void AddProtectedRangeFromSelection()
    {
        if (this.currentSheet is null || this.selectedCellKeys.Count == 0)
        {
            return;
        }

        var parsed = this.selectedCellKeys.Select(key => TryParseCellKey(key, out var row, out var col) ? (row: row, col: col) : (row: 0, col: 0)).Where(x => x.row > 0 && x.col > 0).ToList();
        if (parsed.Count == 0)
        {
            return;
        }

        var settings = this.GetSheetSettings();
        settings.ProtectedRanges.Add(new ProtectedRange
        {
            StartRow = parsed.Min(x => x.row),
            EndRow = parsed.Max(x => x.row),
            StartColumn = parsed.Min(x => x.col),
            EndColumn = parsed.Max(x => x.col),
            AllowOwnerOnly = false,
            EditorUserIds = new List<string> { this.configuration.UserId },
        });
        this.RecordActivity("Protected range added", null, null, null);
    }

    private void RemoveProtectionFromSelection()
    {
        if (this.currentSheet is null || this.selectedCellKeys.Count == 0)
        {
            return;
        }

        var parsed = this.selectedCellKeys.Select(key => TryParseCellKey(key, out var row, out var col) ? (row: row, col: col) : (row: 0, col: 0)).Where(x => x.row > 0 && x.col > 0).ToList();
        if (parsed.Count == 0)
        {
            return;
        }

        var minRow = parsed.Min(x => x.row);
        var maxRow = parsed.Max(x => x.row);
        var minCol = parsed.Min(x => x.col);
        var maxCol = parsed.Max(x => x.col);
        var settings = this.GetSheetSettings();
        settings.ProtectedRanges.RemoveAll(range => range.StartRow == minRow && range.EndRow == maxRow && range.StartColumn == minCol && range.EndColumn == maxCol);
        this.RecordActivity("Protected range removed", null, null, null);
    }

    private void RedoLastChange()
    {
        if (this.TryRedoSharedChange())
        {
            return;
        }

        if (this.currentSheet is null || this.redoHistory.Count == 0)
        {
            return;
        }

        this.undoHistory.Push(SheetSerializationHelper.SerializeForSnapshot(this.currentSheet.Data, this.currentSheet.RowsCount, this.currentSheet.ColsCount));
        var snapshot = this.redoHistory.Pop();
        var restored = JsonSerializer.Deserialize<SheetDocument>(snapshot) ?? SheetDocument.CreateDefault();
        restored.Normalize(this.currentSheet.RowsCount, this.currentSheet.ColsCount);
        this.currentSheet.Data = restored;
        this.statusMessage = "Redid the last undone change.";
    }

    private string GetRenderedCellText(SheetTabData tab, string cellKey)
    {
        var display = this.GetDisplayText(tab, cellKey);
        if (!tab.TryGetFeature(cellKey, out var feature) || feature is null)
        {
            return display;
        }

        if (feature.IsChecklistCell)
        {
            var prefix = feature.IsChecked ? "[x] " : "[ ] ";
            return prefix + display;
        }

        return display;
    }

    private SheetCellData GetEffectiveCellStyle(SheetTabData tab, string cellKey, SheetCellData source)
    {
        var clone = JsonSerializer.Deserialize<SheetCellData>(JsonSerializer.Serialize(source)) ?? new SheetCellData();
        var settings = this.GetSheetSettings();
        var currentValue = this.GetDisplayText(tab, cellKey);
        foreach (var rule in settings.ConditionalRules)
        {
            if (!this.DoesConditionalRuleMatch(rule, currentValue))
            {
                continue;
            }

            if (rule.TextColor.HasValue)
            {
                clone.TextColor = rule.TextColor.Value;
            }

            if (rule.BackgroundColor.HasValue)
            {
                clone.BackgroundColor = rule.BackgroundColor.Value;
            }

            if (rule.Bold.HasValue)
            {
                clone.Bold = rule.Bold.Value;
            }

            if (rule.Italic.HasValue)
            {
                clone.Italic = rule.Italic.Value;
            }

            if (rule.Underline.HasValue)
            {
                clone.Underline = rule.Underline.Value;
            }
        }

        return clone;
    }

    private bool DoesConditionalRuleMatch(ConditionalFormattingRule rule, string currentValue)
    {
        var lhs = currentValue ?? string.Empty;
        var rhs = rule.TargetValue ?? string.Empty;
        var lhsNum = double.TryParse(lhs.Replace(",", string.Empty, StringComparison.Ordinal), NumberStyles.Any, CultureInfo.InvariantCulture, out var lhsParsed) ? lhsParsed : double.NaN;
        var rhsNum = double.TryParse(rhs.Replace(",", string.Empty, StringComparison.Ordinal), NumberStyles.Any, CultureInfo.InvariantCulture, out var rhsParsed) ? rhsParsed : double.NaN;

        return rule.Operator switch
        {
            ConditionalOperator.Equal => string.Equals(lhs, rhs, StringComparison.OrdinalIgnoreCase) || (!double.IsNaN(lhsNum) && !double.IsNaN(rhsNum) && Math.Abs(lhsNum - rhsNum) < 0.000001d),
            ConditionalOperator.NotEqual => !string.Equals(lhs, rhs, StringComparison.OrdinalIgnoreCase),
            ConditionalOperator.GreaterThan => !double.IsNaN(lhsNum) && !double.IsNaN(rhsNum) && lhsNum > rhsNum,
            ConditionalOperator.GreaterThanOrEqual => !double.IsNaN(lhsNum) && !double.IsNaN(rhsNum) && lhsNum >= rhsNum,
            ConditionalOperator.LessThan => !double.IsNaN(lhsNum) && !double.IsNaN(rhsNum) && lhsNum < rhsNum,
            ConditionalOperator.LessThanOrEqual => !double.IsNaN(lhsNum) && !double.IsNaN(rhsNum) && lhsNum <= rhsNum,
            ConditionalOperator.Contains => lhs.Contains(rhs, StringComparison.OrdinalIgnoreCase),
            ConditionalOperator.StartsWith => lhs.StartsWith(rhs, StringComparison.OrdinalIgnoreCase),
            ConditionalOperator.EndsWith => lhs.EndsWith(rhs, StringComparison.OrdinalIgnoreCase),
            _ => false,
        };
    }

    private void EnsurePermissionDefaults(ExpandedSheetSettings settings)
    {
        settings.AutoSaveCloud = true;
        settings.AutoSaveSeconds = 20;
        settings.PermissionPresets ??= new List<SheetPermissionPreset>();
        settings.MemberProfiles ??= new List<SheetMemberProfile>();
        settings.TabPresets ??= new List<SheetTabPreset>();
        settings.TemporaryInviteCodes ??= new List<TemporaryInviteCode>();
        settings.InviteAuditLog ??= new List<InviteAuditEntry>();
        if (!settings.PermissionPresets.Any(p => string.Equals(p.Name, "Viewer", StringComparison.OrdinalIgnoreCase)))
        {
            settings.PermissionPresets.Add(new SheetPermissionPreset
            {
                Name = "Viewer",
                Color = 0xFF5C5C5Cu,
                Permissions = new SheetPermissionSet(),
            });
        }

        if (!settings.PermissionPresets.Any(p => string.Equals(p.Name, "Editor", StringComparison.OrdinalIgnoreCase)))
        {
            var editorPermissions = new SheetPermissionSet
            {
                EditSheet = true,
                CreateTabs = true,
                SeeHistory = true,
                UseComments = true,
                ImportSheet = true,
                SaveLocal = true,
            };
            settings.PermissionPresets.Add(new SheetPermissionPreset
            {
                Name = "Editor",
                Color = 0xFF4F9D4Fu,
                Permissions = editorPermissions,
            });
        }

        if (!settings.PermissionPresets.Any(p => string.Equals(p.Name, "Manager", StringComparison.OrdinalIgnoreCase)))
        {
            var managerPermissions = new SheetPermissionSet
            {
                EditSheet = true,
                EditPermissions = true,
                CreateTabs = true,
                SeeHistory = true,
                UseComments = true,
                ImportSheet = true,
                SaveLocal = true,
                BlockUsers = true,
            };
            settings.PermissionPresets.Add(new SheetPermissionPreset
            {
                Name = "Manager",
                Color = 0xFF2F7FD1u,
                Permissions = managerPermissions,
            });
        }
    }

    private void SyncCurrentUserMemberProfile(SheetDocument document, DateTimeOffset? joinedAtUtc)
    {
        if (string.IsNullOrWhiteSpace(this.configuration.UserId))
        {
            return;
        }

        this.EnsureAdvancedDefaults(document);
        var settings = document.Settings;
        var profile = this.GetOrCreateMemberProfile(settings, this.configuration.UserId, this.GetSafeCharacterFullName(), joinedAtUtc);
        profile.CharacterName = this.GetSafeCharacterFullName();
        profile.LastSeenUtc = DateTimeOffset.UtcNow;
        if (string.IsNullOrWhiteSpace(profile.AssignedPresetName))
        {
            profile.AssignedPresetName = "Viewer";
        }

        if (profile.RoleColor == 0)
        {
            profile.RoleColor = this.FindPermissionPreset(settings, profile.AssignedPresetName)?.Color ?? 0xFF5C5C5Cu;
        }
    }

    private SheetMemberProfile GetOrCreateMemberProfile(ExpandedSheetSettings settings, string userId, string fallbackName, DateTimeOffset? joinedAtUtc = null)
    {
        this.EnsurePermissionDefaults(settings);
        var profile = settings.MemberProfiles.FirstOrDefault(x => string.Equals(x.UserId, userId, StringComparison.OrdinalIgnoreCase));
        if (profile is not null)
        {
            if (!string.IsNullOrWhiteSpace(fallbackName) && string.IsNullOrWhiteSpace(profile.CharacterName))
            {
                profile.CharacterName = fallbackName;
            }

            if (!profile.JoinedAtUtc.HasValue && joinedAtUtc.HasValue)
            {
                profile.JoinedAtUtc = joinedAtUtc.Value;
            }

            if (string.IsNullOrWhiteSpace(profile.AssignedPresetName))
            {
                profile.AssignedPresetName = "Viewer";
            }

            return profile;
        }

        var preset = this.FindPermissionPreset(settings, "Viewer");
        profile = new SheetMemberProfile
        {
            UserId = userId,
            CharacterName = fallbackName,
            JoinedAtUtc = joinedAtUtc ?? DateTimeOffset.UtcNow,
            LastSeenUtc = DateTimeOffset.UtcNow,
            AssignedPresetName = preset?.Name ?? "Viewer",
            RoleColor = preset?.Color ?? 0xFF5C5C5Cu,
            Permissions = new SheetPermissionSet(),
        };
        if (preset is not null)
        {
            profile.Permissions.CopyFrom(preset.Permissions);
        }

        settings.MemberProfiles.Add(profile);
        return profile;
    }

    private SheetMemberProfile? FindMemberProfile(ExpandedSheetSettings settings, string userId)
        => settings.MemberProfiles.FirstOrDefault(x => string.Equals(x.UserId, userId, StringComparison.OrdinalIgnoreCase));

    private SheetPermissionPreset? FindPermissionPreset(ExpandedSheetSettings settings, string? presetName)
    {
        if (string.IsNullOrWhiteSpace(presetName))
        {
            return null;
        }

        return settings.PermissionPresets.FirstOrDefault(x => string.Equals(x.Name, presetName, StringComparison.OrdinalIgnoreCase));
    }

    private void ApplyPresetToProfile(SheetMemberProfile profile, SheetPermissionPreset preset)
    {
        profile.AssignedPresetName = preset.Name;
        profile.RoleColor = preset.Color;
        profile.IsBlocked = false;
        profile.Permissions.CopyFrom(preset.Permissions);
    }

    private void MarkMemberProfileBlocked(SheetMemberProfile profile, bool blocked)
    {
        profile.IsBlocked = blocked;
        if (blocked)
        {
            profile.AssignedPresetName = "Blocked";
            profile.RoleColor = 0xFFC84343u;
            profile.Permissions.Reset();
        }
        else if (string.Equals(profile.AssignedPresetName, "Blocked", StringComparison.OrdinalIgnoreCase))
        {
            profile.AssignedPresetName = "Viewer";
            profile.RoleColor = this.FindPermissionPreset(this.GetSheetSettings(), "Viewer")?.Color ?? 0xFF5C5C5Cu;
        }
    }

    private bool HasSheetPermission(SheetPermissionType permission)
    {
        if (this.currentSheet is null)
        {
            return false;
        }

        if (this.currentRole == SheetAccessRole.Owner)
        {
            return true;
        }

        var settings = this.GetSheetSettings();
        this.EnsurePermissionDefaults(settings);
        var profile = this.FindMemberProfile(settings, this.configuration.UserId);
        if (profile is null || profile.IsBlocked)
        {
            return false;
        }

        return profile.Permissions.Has(permission);
    }

    private bool CanOpenPermissionsPanel()
        => this.currentSheet is not null
           && (this.HasSheetPermission(SheetPermissionType.EditPermissions)
               || this.HasSheetPermission(SheetPermissionType.BlockUsers));

    private bool CanUseInviteCodes()
        => this.currentSheet is not null && (this.currentRole == SheetAccessRole.Owner || this.HasSheetPermission(SheetPermissionType.Invite));

    private bool CanRemoveMembers()
        => this.currentSheet is not null && (this.currentRole == SheetAccessRole.Owner || this.HasSheetPermission(SheetPermissionType.BlockUsers));

    private bool CanManageRoles()
        => this.currentSheet is not null && (this.currentRole == SheetAccessRole.Owner || this.HasSheetPermission(SheetPermissionType.EditPermissions));

    private bool CanEditOwnRole(string targetUserId)
        => this.currentSheet is not null
           && string.Equals(targetUserId, this.configuration.UserId, StringComparison.OrdinalIgnoreCase)
           && this.currentRole != SheetAccessRole.Owner;

    private bool CanDeleteCurrentSheet()
        => this.currentSheet is not null && (this.currentRole == SheetAccessRole.Owner || this.HasSheetPermission(SheetPermissionType.DeleteSheet));

    private bool CanCreateTabs()
        => this.currentSheet is not null
           && this.currentSheet.Data.Settings.ViewMode != SheetViewMode.ReadOnly
           && (this.currentRole == SheetAccessRole.Owner
               || this.HasSheetPermission(SheetPermissionType.CreateTabs)
               || this.HasSheetPermission(SheetPermissionType.EditSheet));

    private bool CanPersistCurrentSheetChanges()
        => this.currentSheet is not null
           && this.currentSheet.Data.Settings.ViewMode != SheetViewMode.ReadOnly
           && (this.currentRole == SheetAccessRole.Owner
               || this.HasSheetPermission(SheetPermissionType.EditSheet)
               || this.HasSheetPermission(SheetPermissionType.EditPermissions)
               || this.HasSheetPermission(SheetPermissionType.CreateTabs)
               || this.HasSheetPermission(SheetPermissionType.UseComments)
               || this.HasSheetPermission(SheetPermissionType.SeeHistory)
               || this.HasSheetPermission(SheetPermissionType.ImportSheet)
               || this.HasSheetPermission(SheetPermissionType.Admin));

    private string GetDisplayedRoleLabel()
    {
        if (this.currentSheet is null)
        {
            return this.currentRole.ToString();
        }

        if (this.currentRole == SheetAccessRole.Owner)
        {
            return "Owner";
        }

        var settings = this.GetSheetSettings();
        var profile = this.FindMemberProfile(settings, this.configuration.UserId);
        if (profile is null)
        {
            return this.currentRole.ToString();
        }

        if (profile.IsBlocked)
        {
            return "Blocked";
        }

        return string.IsNullOrWhiteSpace(profile.AssignedPresetName) ? this.currentRole.ToString() : profile.AssignedPresetName;
    }

    private Vector4 GetDisplayedRoleColor()
    {
        if (this.currentRole == SheetAccessRole.Owner)
        {
            return HexToVector4("#d99a32");
        }

        var settings = this.GetSheetSettings();
        var profile = this.FindMemberProfile(settings, this.configuration.UserId);
        if (profile is null)
        {
            return GrayDisabled;
        }

        return ImGui.ColorConvertU32ToFloat4(profile.RoleColor == 0 ? 0xFF9A9A9Au : profile.RoleColor);
    }

    private string GetDisplayedRoleTooltip()
    {
        if (this.currentRole == SheetAccessRole.Owner)
        {
            return "Owner has full access to this sheet.";
        }

        var settings = this.GetSheetSettings();
        var profile = this.FindMemberProfile(settings, this.configuration.UserId);
        if (profile is null)
        {
            return "No preset has been assigned yet.";
        }

        if (profile.IsBlocked)
        {
            return "This role is blocked and has no active permissions.";
        }

        var labels = profile.Permissions.EnumerateAllowedLabels().ToList();
        return labels.Count == 0 ? "This role currently has no active permissions." : string.Join("\n", labels);
    }

    private string GetPermissionSummary(SheetPermissionSet permissions, bool isBlocked)
    {
        if (isBlocked)
        {
            return "Blocked";
        }

        var labels = permissions.EnumerateAllowedLabels().ToList();
        if (labels.Count == 0)
        {
            return "No access";
        }

        if (labels.Count <= 2)
        {
            return string.Join(", ", labels);
        }

        return $"{labels.Count} enabled";
    }

    private async Task RefreshSheetMembersAsync(CancellationToken cancellationToken = default)
    {
        if (this.currentSheet is null)
        {
            lock (this.currentSheetMembers)
            {
                this.currentSheetMembers.Clear();
            }
            this.lastSheetMembersRefreshUtc = DateTimeOffset.UtcNow;
            return;
        }

        var settings = this.GetSheetSettings();
        var rows = await this.supabase.ListSheetMembersAsync(this.currentSheet.Id, cancellationToken).ConfigureAwait(false);
        lock (this.currentSheetMembers)
        {
            this.currentSheetMembers.Clear();
            this.currentSheetMembers.AddRange(rows.OrderBy(x => x.CreatedAt).ThenBy(x => x.UserId, StringComparer.OrdinalIgnoreCase));
        }
        this.lastSheetMembersRefreshUtc = DateTimeOffset.UtcNow;

        foreach (var row in rows)
        {
            var fallbackName = !string.IsNullOrWhiteSpace(row.CharacterName) ? row.CharacterName : this.ResolveMemberCharacterName(settings, row.UserId);
            var profile = this.GetOrCreateMemberProfile(settings, row.UserId, fallbackName, row.CreatedAt);
            profile.CharacterName = !string.IsNullOrWhiteSpace(row.CharacterName) ? row.CharacterName : profile.CharacterName;
            profile.LastSeenUtc = row.LastSeenUtc ?? profile.LastSeenUtc;
            if (!string.IsNullOrWhiteSpace(row.AssignedPresetName))
            {
                profile.AssignedPresetName = row.AssignedPresetName;
            }
            if (row.RoleColor != 0)
            {
                profile.RoleColor = row.RoleColor;
            }
            profile.IsBlocked = row.IsBlocked;
        }

        this.GetOrCreateMemberProfile(settings, this.currentSheet.OwnerId, this.ResolveMemberCharacterName(settings, this.currentSheet.OwnerId), this.currentSheet.CreatedAt);
    }

    private string ResolveMemberCharacterName(ExpandedSheetSettings settings, string userId)
    {
        if (string.Equals(userId, this.configuration.UserId, StringComparison.OrdinalIgnoreCase))
        {
            return this.GetSafeCharacterFullName();
        }

        SheetMemberRow? memberRow;
        lock (this.currentSheetMembers)
        {
            memberRow = this.currentSheetMembers.FirstOrDefault(x => string.Equals(x.UserId, userId, StringComparison.OrdinalIgnoreCase));
        }
        if (memberRow is not null && !string.IsNullOrWhiteSpace(memberRow.CharacterName))
        {
            return memberRow.CharacterName;
        }

        var profile = this.FindMemberProfile(settings, userId);
        if (profile is not null && !string.IsNullOrWhiteSpace(profile.CharacterName))
        {
            return profile.CharacterName;
        }

        return $"User {userId[..Math.Min(8, userId.Length)]}";
    }

    private async Task RefreshBlocklistAsync(CancellationToken cancellationToken = default)
    {
        if (this.currentSheet is null)
        {
            lock (this.currentSheetBlocklist)
            {
                this.currentSheetBlocklist.Clear();
            }
            this.lastBlocklistRefreshUtc = DateTimeOffset.UtcNow;
            return;
        }

        var rows = await this.supabase.ListSheetBlocklistAsync(this.currentSheet.Id, cancellationToken).ConfigureAwait(false);
        lock (this.currentSheetBlocklist)
        {
            this.currentSheetBlocklist.Clear();
            this.currentSheetBlocklist.AddRange(rows.OrderByDescending(x => x.RemovedAt).ThenBy(x => x.UserId, StringComparer.OrdinalIgnoreCase));
        }
        this.lastBlocklistRefreshUtc = DateTimeOffset.UtcNow;
    }


    private void BeginRefreshSheetMembersInBackground(bool force = false)
    {
        if (this.currentSheet is null || this.sheetMembersRefreshInFlight)
        {
            return;
        }

        if (!force && DateTimeOffset.UtcNow - this.lastSheetMembersRefreshUtc < TimeSpan.FromSeconds(10))
        {
            return;
        }

        this.sheetMembersRefreshInFlight = true;
        _ = Task.Run(async () =>
        {
            try
            {
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(4));
                await this.RefreshSheetMembersAsync(cts.Token).ConfigureAwait(false);
            }
            catch
            {
            }
            finally
            {
                this.sheetMembersRefreshInFlight = false;
            }
        });
    }

    private void BeginRefreshBlocklistInBackground(bool force = false)
    {
        if (this.currentSheet is null || this.blocklistRefreshInFlight)
        {
            return;
        }

        if (!force && DateTimeOffset.UtcNow - this.lastBlocklistRefreshUtc < TimeSpan.FromSeconds(10))
        {
            return;
        }

        this.blocklistRefreshInFlight = true;
        _ = Task.Run(async () =>
        {
            try
            {
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(4));
                await this.RefreshBlocklistAsync(cts.Token).ConfigureAwait(false);
            }
            catch
            {
            }
            finally
            {
                this.blocklistRefreshInFlight = false;
            }
        });
    }

    private void OpenBlocklistWindow()
    {
        this.blocklistWindowOpen = true;
        this.BeginRefreshBlocklistInBackground();
    }

    private void OpenPermissionsWindow()
    {
        this.permissionsWindowOpen = true;
        this.BeginRefreshSheetMembersInBackground();
    }

    private void OpenPermissionsEditorForUser(string? userId)
    {
        if (this.currentSheet is null || string.IsNullOrWhiteSpace(userId) || !this.CanManageRoles())
        {
            return;
        }

        var targetUserId = userId!;
        var settings = this.GetSheetSettings();
        if (this.FindMemberProfile(settings, targetUserId) is null)
        {
            this.BeginRefreshSheetMembersInBackground(force: true);
        }

        this.permissionsWindowOpen = true;
        this.permissionsEditorUserId = targetUserId;
    }

    private void DrawPermissionsWindow()
    {
        if (!this.permissionsWindowOpen || this.currentSheet is null)
        {
            return;
        }

        var open = this.permissionsWindowOpen;
        ImGui.SetNextWindowSize(new Vector2(960f, 430f), ImGuiCond.FirstUseEver);
        if (!ImGui.Begin("Sheet Permissions###EZSheetsPermissions", ref open))
        {
            ImGui.End();
            this.permissionsWindowOpen = open;
            return;
        }

        this.permissionsWindowOpen = open;
        var settings = this.GetSheetSettings();
        this.EnsurePermissionDefaults(settings);
        ImGui.TextWrapped("Owner and authorized managers can assign presets, inspect permissions, and remove access from this sheet.");
        ImGui.Spacing();

        if (this.DrawStyledButton("permissions-refresh-members", "Refresh Users", AccentBlue, WhiteText, new Vector2(112f, 22f)))
        {
            this.OpenPermissionsWindow();
        }

        ImGui.Spacing();
        var tableFlags = ImGuiTableFlags.Borders | ImGuiTableFlags.RowBg | ImGuiTableFlags.ScrollY | ImGuiTableFlags.Resizable | ImGuiTableFlags.SizingStretchProp;
        if (ImGui.BeginTable("##permissions-table", 6, tableFlags, new Vector2(-1f, 320f)))
        {
            ImGui.TableSetupColumn("Name", ImGuiTableColumnFlags.WidthStretch, 2.2f);
            ImGui.TableSetupColumn("Joined", ImGuiTableColumnFlags.WidthFixed, 128f);
            ImGui.TableSetupColumn("Last Seen", ImGuiTableColumnFlags.WidthFixed, 128f);
            ImGui.TableSetupColumn("Role", ImGuiTableColumnFlags.WidthStretch, 1.1f);
            ImGui.TableSetupColumn("Permission", ImGuiTableColumnFlags.WidthStretch, 1.4f);
            ImGui.TableSetupColumn("Actions", ImGuiTableColumnFlags.WidthFixed, 122f);
            ImGui.TableHeadersRow();

            var rowIndex = 0;
            var ownerProfile = this.GetOrCreateMemberProfile(settings, this.currentSheet.OwnerId, this.ResolveMemberCharacterName(settings, this.currentSheet.OwnerId), this.currentSheet.CreatedAt);
            this.DrawPermissionMemberRow(ownerProfile, true, this.currentSheet.CreatedAt, ownerProfile.LastSeenUtc, rowIndex++);
            List<SheetMemberRow> memberSnapshot;
            lock (this.currentSheetMembers)
            {
                memberSnapshot = this.currentSheetMembers.OrderBy(x => x.CreatedAt).ThenBy(x => x.UserId, StringComparer.OrdinalIgnoreCase).ToList();
            }
            foreach (var member in memberSnapshot)
            {
                if (string.Equals(member.UserId, this.currentSheet.OwnerId, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var fallbackName = !string.IsNullOrWhiteSpace(member.CharacterName) ? member.CharacterName : this.ResolveMemberCharacterName(settings, member.UserId);
                var profile = this.GetOrCreateMemberProfile(settings, member.UserId, fallbackName, member.CreatedAt);
                if (!string.IsNullOrWhiteSpace(member.CharacterName))
                {
                    profile.CharacterName = member.CharacterName;
                }
                profile.LastSeenUtc = member.LastSeenUtc ?? profile.LastSeenUtc;
                profile.IsBlocked = member.IsBlocked;
                this.DrawPermissionMemberRow(profile, false, member.CreatedAt, member.LastSeenUtc ?? profile.LastSeenUtc, rowIndex++);
            }

            ImGui.EndTable();
        }

        ImGui.End();
    }

    private void DrawPermissionMemberRow(SheetMemberProfile profile, bool isOwner, DateTimeOffset? joinedAtUtc, DateTimeOffset? lastSeenUtc, int rowIndex)
    {
        var settings = this.GetSheetSettings();
        ImGui.TableNextRow();
        var rowColor = rowIndex % 2 == 0 ? 0x44363636u : 0x44262626u;
        ImGui.TableSetBgColor(ImGuiTableBgTarget.RowBg0, rowColor);
        ImGui.TableNextColumn();
        ImGui.TextUnformatted(string.IsNullOrWhiteSpace(profile.CharacterName) ? profile.UserId : profile.CharacterName);
        ImGui.TableNextColumn();
        ImGui.TextUnformatted(joinedAtUtc?.LocalDateTime.ToString("g") ?? "-");
        ImGui.TableNextColumn();
        ImGui.TextUnformatted(lastSeenUtc?.LocalDateTime.ToString("g") ?? "-");
        ImGui.TableNextColumn();
        if (isOwner)
        {
            ImGui.TextColored(HexToVector4("#d99a32"), "Owner");
        }
        else
        {
            var roleText = string.IsNullOrWhiteSpace(profile.AssignedPresetName) ? "Viewer" : profile.AssignedPresetName;
            var textColor = profile.RoleColor != 0 ? ImGui.ColorConvertU32ToFloat4(profile.RoleColor) : HexToVector4("#5c5c5c");
            var label = $" {roleText} ";
            var labelSize = ImGui.CalcTextSize(label);
            var labelMin = ImGui.GetCursorScreenPos();
            var labelMax = labelMin + new Vector2(labelSize.X + 8f, 20f);
            var drawList = ImGui.GetWindowDrawList();
            drawList.AddRectFilled(labelMin, labelMax, ImGui.ColorConvertFloat4ToU32(new Vector4(1f, 1f, 1f, 0.18f)), 3f);
            drawList.AddText(labelMin + new Vector2(4f, 2f), ImGui.ColorConvertFloat4ToU32(textColor), roleText);
            ImGui.InvisibleButton($"perm-role-label-{profile.UserId}", new Vector2(labelSize.X + 8f, 20f));
            ImGui.SameLine(0f, 4f);
            var canChangeTargetRole = this.CanManageRoles() && !this.CanEditOwnRole(profile.UserId);
            ImGui.BeginDisabled(!canChangeTargetRole);
            if (this.DrawStyledButton($"perm-role-picker-{profile.UserId}", "◎", HexToVector4("#4f4f4f"), WhiteText, new Vector2(22f, 20f)) && canChangeTargetRole)
            {
                ImGui.OpenPopup($"perm-role-popup-{profile.UserId}");
            }
            ImGui.EndDisabled();
            if (ImGui.BeginPopup($"perm-role-popup-{profile.UserId}"))
            {
                foreach (var preset in settings.PermissionPresets.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase))
                {
                    if (ImGui.Selectable(preset.Name) && canChangeTargetRole)
                    {
                        this.ApplyPresetToProfile(profile, preset);
                        ImGui.CloseCurrentPopup();
                        this.RequestBackgroundCloudSave("Could not save permission changes.");
                    }
                }
                ImGui.EndPopup();
            }
        }

        ImGui.TableNextColumn();
        var summary = this.GetPermissionSummary(profile.Permissions, profile.IsBlocked);
        if (profile.AccessExpiresAtUtc.HasValue)
        {
            summary = string.IsNullOrWhiteSpace(summary)
                ? $"Temp until {profile.AccessExpiresAtUtc.Value.LocalDateTime:g}"
                : summary + "\n" + $"Temp until {profile.AccessExpiresAtUtc.Value.LocalDateTime:g}";
        }
        ImGui.TextWrapped(summary);
        ImGui.TableNextColumn();
        var canOpenEditorForTarget = !isOwner && this.CanManageRoles() && !this.CanEditOwnRole(profile.UserId);
        ImGui.BeginDisabled(!canOpenEditorForTarget);
        if (this.DrawStyledButton($"perm-open-{profile.UserId}", "P", HexToVector4("#4f76d1"), WhiteText, new Vector2(24f, 20f)) && canOpenEditorForTarget)
        {
            this.permissionsEditorUserId = profile.UserId;
            this.permissionPresetDraft = profile.AssignedPresetName;
            this.permissionPresetColor = ImGui.ColorConvertU32ToFloat4(profile.RoleColor == 0 ? 0xFFD46728u : profile.RoleColor);
        }
        ImGui.EndDisabled();

        if (!isOwner)
        {
            ImGui.SameLine(0, 4f);
            var canRemoveTarget = this.CanRemoveMembers() && !string.Equals(profile.UserId, this.configuration.UserId, StringComparison.OrdinalIgnoreCase);
            var removeColor = canRemoveTarget ? HexToVector4("#af3f3f") : HexToVector4("#565656");
            if (this.DrawStyledButton($"perm-remove-{profile.UserId}", "X", removeColor, WhiteText, new Vector2(24f, 20f)) && canRemoveTarget)
            {
                this.pendingRemovalReasonUserId = profile.UserId;
                this.pendingRemovalReasonText = string.Empty;
                ImGui.OpenPopup($"perm-remove-popup-{profile.UserId}");
            }
            if (ImGui.IsItemHovered() && !canRemoveTarget)
            {
                ImGui.SetTooltip("You dont have permission to remove user");
            }
            if (ImGui.BeginPopup($"perm-remove-popup-{profile.UserId}"))
            {
                ImGui.TextUnformatted("Reason:");
                ImGui.SetNextItemWidth(180f);
                ImGui.InputText($"##perm-remove-reason-{profile.UserId}", ref this.pendingRemovalReasonText, 128);
                ImGui.SameLine(0, 6f);
                if (this.DrawStyledButton($"perm-remove-confirm-{profile.UserId}", "Remove", AccentDelete, WhiteText, new Vector2(72f, 22f)))
                {
                    var reason = this.pendingRemovalReasonText.Trim();
                    this.pendingRemovalReasonUserId = string.Empty;
                    this.pendingRemovalReasonText = string.Empty;
                    ImGui.CloseCurrentPopup();
                    _ = this.RunActionAsync(async () =>
                    {
                        if (this.currentSheet is null)
                        {
                            return;
                        }

                        await this.supabase.RemoveSheetMemberAsync(this.currentSheet.Id, profile.UserId, reason).ConfigureAwait(false);
                        await this.RefreshSheetMembersAsync().ConfigureAwait(false);
                        await this.RefreshBlocklistAsync().ConfigureAwait(false);
                        await this.RefreshLiveSheetStateAsync(this.currentSheet.Id).ConfigureAwait(false);
                        this.memberRemovedName = string.IsNullOrWhiteSpace(profile.CharacterName) ? "User" : profile.CharacterName;
                        this.memberRemovedPopupOpen = true;
                    }, "Could not remove that sheet member.");
                }
                ImGui.EndPopup();
            }

        }
    }

    private void DrawBlocklistWindow()
    {
        if (!this.blocklistWindowOpen || this.currentSheet is null)
        {
            return;
        }

        var open = this.blocklistWindowOpen;
        ImGui.SetNextWindowSize(new Vector2(720f, 320f), ImGuiCond.FirstUseEver);
        if (!ImGui.Begin("Sheet Blocklist###EZSheetsBlocklist", ref open))
        {
            ImGui.End();
            this.blocklistWindowOpen = open;
            return;
        }

        this.blocklistWindowOpen = open;
        if (this.DrawStyledButton("blocklist-refresh", "Refresh", AccentBlue, WhiteText, new Vector2(92f, 22f)))
        {
            this.OpenBlocklistWindow();
        }

        var flags = ImGuiTableFlags.Borders | ImGuiTableFlags.RowBg | ImGuiTableFlags.ScrollY | ImGuiTableFlags.Resizable | ImGuiTableFlags.SizingStretchProp;
        if (ImGui.BeginTable("##blocklist-table", 4, flags, new Vector2(-1f, 220f)))
        {
            ImGui.TableSetupColumn("Name", ImGuiTableColumnFlags.WidthStretch, 1.4f);
            ImGui.TableSetupColumn("Reason", ImGuiTableColumnFlags.WidthStretch, 2.2f);
            ImGui.TableSetupColumn("Removal Date", ImGuiTableColumnFlags.WidthFixed, 150f);
            ImGui.TableSetupColumn("Actions", ImGuiTableColumnFlags.WidthFixed, 90f);
            ImGui.TableHeadersRow();

            List<SheetBlockedMemberRow> snapshot;
            lock (this.currentSheetBlocklist)
            {
                snapshot = this.currentSheetBlocklist.ToList();
            }

            var rowIndex = 0;
            foreach (var entry in snapshot)
            {
                ImGui.TableNextRow();
                ImGui.TableSetBgColor(ImGuiTableBgTarget.RowBg0, rowIndex++ % 2 == 0 ? 0x44363636u : 0x44262626u);
                ImGui.TableNextColumn();
                ImGui.TextUnformatted(string.IsNullOrWhiteSpace(entry.CharacterName) ? $"User {entry.UserId[..Math.Min(8, entry.UserId.Length)]}" : entry.CharacterName);
                ImGui.TableNextColumn();
                ImGui.TextWrapped(string.IsNullOrWhiteSpace(entry.Reason) ? "-" : entry.Reason);
                ImGui.TableNextColumn();
                ImGui.TextUnformatted(entry.RemovedAt.LocalDateTime.ToString("g"));
                ImGui.TableNextColumn();
                if (this.DrawStyledButton($"blocklist-unblock-{entry.UserId}", "Unblock", AccentDelete, WhiteText, new Vector2(72f, 22f)))
                {
                    _ = this.RunActionAsync(async () =>
                    {
                        if (this.currentSheet is null)
                        {
                            return;
                        }

                        await this.supabase.UnblockSheetMemberAsync(this.currentSheet.Id, entry.UserId).ConfigureAwait(false);
                        await this.RefreshBlocklistAsync().ConfigureAwait(false);
                        await this.RefreshSheetMembersAsync().ConfigureAwait(false);
                    }, "Could not unblock that user.");
                }
            }

            ImGui.EndTable();
        }

        ImGui.End();
    }

    private void DrawPermissionEditorWindow()
    {
        if (this.currentSheet is null || string.IsNullOrWhiteSpace(this.permissionsEditorUserId))
        {
            return;
        }

        var settings = this.GetSheetSettings();
        var profile = this.FindMemberProfile(settings, this.permissionsEditorUserId);
        if (profile is null)
        {
            this.permissionsEditorUserId = string.Empty;
            return;
        }

        var open = true;
        ImGui.SetNextWindowSize(new Vector2(460f, 520f), ImGuiCond.FirstUseEver);
        if (!ImGui.Begin($"Permissions for {profile.CharacterName}###EZSheetsPermissionEditor", ref open, ImGuiWindowFlags.NoCollapse))
        {
            ImGui.End();
            if (!open)
            {
                this.permissionsEditorUserId = string.Empty;
            }
            return;
        }

        if (!open)
        {
            ImGui.End();
            this.permissionsEditorUserId = string.Empty;
            return;
        }

        ImGui.TextWrapped(profile.CharacterName);
        ImGui.Spacing();

        var canManageTarget = !this.CanEditOwnRole(profile.UserId) && this.CanManageRoles();
        var canSavePresets = this.currentRole == SheetAccessRole.Owner;
        var presetNames = settings.PermissionPresets.Select(x => x.Name).ToArray();
        var selectedPresetIndex = Math.Max(0, Array.FindIndex(presetNames, x => string.Equals(x, this.permissionPresetDraft, StringComparison.OrdinalIgnoreCase)));
        if (presetNames.Length == 0)
        {
            presetNames = new[] { "Viewer" };
            selectedPresetIndex = 0;
        }

        ImGui.SetNextItemWidth(220f);
        ImGui.BeginDisabled(!canManageTarget);
        if (ImGui.Combo("Preset", ref selectedPresetIndex, presetNames, presetNames.Length))
        {
            var selectedName = presetNames[Math.Clamp(selectedPresetIndex, 0, presetNames.Length - 1)];
            var preset = this.FindPermissionPreset(settings, selectedName);
            if (preset is not null)
            {
                this.permissionPresetDraft = preset.Name;
                this.permissionPresetColor = ImGui.ColorConvertU32ToFloat4(preset.Color == 0 ? 0xFF5C5C5Cu : preset.Color);
                profile.Permissions.CopyFrom(preset.Permissions);
            }
        }
        ImGui.EndDisabled();

        ImGui.SameLine(0f, 6f);
        ImGui.BeginDisabled(!canManageTarget);
        if (this.DrawStyledButton("permissions-give-role", "Give Role", AccentBlue, WhiteText, new Vector2(92f, 22f)) && canManageTarget)
        {
            var trimmed = this.permissionPresetDraft.Trim();
            if (!string.IsNullOrWhiteSpace(trimmed))
            {
                profile.AssignedPresetName = trimmed;
                profile.RoleColor = ImGui.ColorConvertFloat4ToU32(this.permissionPresetColor);
                this.RequestBackgroundCloudSave("Could not save permission changes.");
            }
        }
        ImGui.EndDisabled();

        if (canSavePresets)
        {
            var presetColor = this.permissionPresetColor;
            ImGui.SetNextItemWidth(120f);
            if (ImGui.ColorEdit4("Preset color", ref presetColor, ImGuiColorEditFlags.NoInputs | ImGuiColorEditFlags.AlphaBar))
            {
                this.permissionPresetColor = presetColor;
            }

            ImGui.SetNextItemWidth(220f);
            ImGui.InputText("Preset name", ref this.permissionPresetDraft, 64);
            if (this.DrawStyledButton("permissions-save-preset", "Save preset", AccentGreen, WhiteText, new Vector2(110f, 22f)))
            {
                var trimmed = this.permissionPresetDraft.Trim();
                if (!string.IsNullOrWhiteSpace(trimmed))
                {
                    var existingPreset = this.FindPermissionPreset(settings, trimmed);
                    if (existingPreset is null)
                    {
                        existingPreset = new SheetPermissionPreset { Name = trimmed };
                        settings.PermissionPresets.Add(existingPreset);
                    }

                    existingPreset.Color = ImGui.ColorConvertFloat4ToU32(this.permissionPresetColor);
                    existingPreset.Permissions.CopyFrom(profile.Permissions);
                    this.RequestBackgroundCloudSave("Could not save permission changes.");
                }
            }
        }

        ImGui.Spacing();
        ImGui.Spacing();
        this.DrawPermissionToggleRow(profile, SheetPermissionType.EditSheet, "Edit Sheet");
        this.DrawPermissionToggleRow(profile, SheetPermissionType.DeleteSheet, "Delete Sheet");
        this.DrawPermissionToggleRow(profile, SheetPermissionType.EditPermissions, "Edit Role");
        this.DrawPermissionToggleRow(profile, SheetPermissionType.CreateTabs, "Create Tabs");
        this.DrawPermissionToggleRow(profile, SheetPermissionType.SeeHistory, "See History");
        this.DrawPermissionToggleRow(profile, SheetPermissionType.UseComments, "Use Comments");
        this.DrawPermissionToggleRow(profile, SheetPermissionType.ImportSheet, "Import Sheet");
        this.DrawPermissionToggleRow(profile, SheetPermissionType.SaveLocal, "Save Local");
        this.DrawPermissionToggleRow(profile, SheetPermissionType.Invite, "Invite");
        this.DrawPermissionToggleRow(profile, SheetPermissionType.BlockUsers, "Block Users");
        this.DrawPermissionToggleRow(profile, SheetPermissionType.Admin, "Admin");

        if (this.DrawStyledButton("permissions-close-editor", "Close", HexToVector4("#555555"), WhiteText, new Vector2(92f, 22f)))
        {
            this.permissionsEditorUserId = string.Empty;
        }

        ImGui.End();
    }

    private void DrawPermissionToggleRow(SheetMemberProfile profile, SheetPermissionType permission, string label)
    {
        var allowed = profile.Permissions.Has(permission);
        var denied = profile.IsBlocked || !allowed;
        var canManageTarget = this.CanManageRoles() && !this.CanEditOwnRole(profile.UserId);
        ImGui.AlignTextToFramePadding();
        ImGui.TextUnformatted(label);
        ImGui.SameLine(220f);
        var allowColor = allowed ? HexToVector4("#2f8f45") : HexToVector4("#4b4b4b");
        ImGui.BeginDisabled(!canManageTarget);
        if (this.DrawStyledButton($"perm-allow-{profile.UserId}-{permission}", "✓", allowColor, WhiteText, new Vector2(28f, 20f)) && canManageTarget)
        {
            profile.IsBlocked = false;
            profile.Permissions.GetType().GetProperty(permission.ToString())?.SetValue(profile.Permissions, true);
            if (!string.Equals(profile.AssignedPresetName, "Custom", StringComparison.OrdinalIgnoreCase))
            {
                profile.AssignedPresetName = "Custom";
            }
            profile.RoleColor = profile.RoleColor == 0 ? 0xFFD46728u : profile.RoleColor;
            this.RequestBackgroundCloudSave("Could not save permission changes.");
        }

        ImGui.SameLine(0, 6f);
        var denyColor = denied ? HexToVector4("#a33b3b") : HexToVector4("#4b4b4b");
        if (this.DrawStyledButton($"perm-deny-{profile.UserId}-{permission}", "X", denyColor, WhiteText, new Vector2(28f, 20f)) && canManageTarget)
        {
            profile.Permissions.GetType().GetProperty(permission.ToString())?.SetValue(profile.Permissions, false);
            if (permission == SheetPermissionType.Admin)
            {
                profile.Permissions.Admin = false;
            }
            if (!string.Equals(profile.AssignedPresetName, "Custom", StringComparison.OrdinalIgnoreCase))
            {
                profile.AssignedPresetName = "Custom";
            }
            this.RequestBackgroundCloudSave("Could not save permission changes.");
        }
        ImGui.EndDisabled();
    }


    private string ComputePersistedSheetFingerprint()
    {
        if (this.currentSheet is null)
        {
            return string.Empty;
        }

        var clone = SheetSerializationHelper.CloneForPersistence(this.currentSheet.Data, this.currentSheet.RowsCount, this.currentSheet.ColsCount, aggressiveTrim: false);

        return JsonSerializer.Serialize(new
        {
            this.currentSheet.Title,
            this.currentSheet.RowsCount,
            this.currentSheet.ColsCount,
            this.currentSheet.DefaultRole,
            Data = clone,
        });
    }

    private void SyncPersistedSheetFingerprint()
    {
        this.lastPersistedSheetFingerprint = string.Empty;
        this.persistableSheetDirty = false;
        this.lastPersistableChangeUtc = DateTimeOffset.MinValue;
    }

    private void MarkPersistableSheetDirty()
    {
        if (this.currentSheet is null)
        {
            return;
        }

        this.persistableSheetDirty = true;
        this.lastPersistableChangeUtc = DateTimeOffset.UtcNow;
    }

    private void RequestBackgroundCloudSave(string fallbackError, int debounceMilliseconds = 1200)
    {
        this.MarkPersistableSheetDirty();

        if (this.currentSheet is null || !this.supabase.IsAuthenticated || !this.CanPersistCurrentSheetChanges())
        {
            return;
        }

        var requestId = System.Threading.Interlocked.Increment(ref this.pendingBackgroundCloudSaveRequestId);
        _ = Task.Run(async () =>
        {
            try
            {
                await Task.Delay(Math.Max(150, debounceMilliseconds)).ConfigureAwait(false);
                if (requestId != this.pendingBackgroundCloudSaveRequestId)
                {
                    return;
                }

                if (this.currentSheet is null || !this.persistableSheetDirty || this.actionInProgress)
                {
                    return;
                }

                this.suppressNextCloudSaveHistory = true;
                this.suppressNextCloudSaveStatus = true;
                await this.SaveCurrentSheetToCloudAsync().ConfigureAwait(false);
                this.autoSaveBackoffUntilUtc = DateTimeOffset.MinValue;
            }
            catch (Exception ex)
            {
                var message = string.IsNullOrWhiteSpace(ex.Message) ? fallbackError : ex.Message;
                this.statusMessage = message;
                Plugin.Log.Warning(ex, fallbackError);
            }
        });
    }

    private bool HasPersistableChangesSinceLastSave()
    {
        return this.currentSheet is not null && this.persistableSheetDirty;
    }

    private void CreateTabFromSavedPreset(SheetTabPreset preset)
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var document = this.currentSheet.Data;
        document.Normalize(this.currentSheet.RowsCount, this.currentSheet.ColsCount);
        var clonedTab = JsonSerializer.Deserialize<SheetTabData>(JsonSerializer.Serialize(preset.Tab)) ?? new SheetTabData();
        clonedTab.Cells ??= new Dictionary<string, SheetCellData>(StringComparer.OrdinalIgnoreCase);
        clonedTab.CellFeatures ??= new Dictionary<string, CellFeatureBundle>(StringComparer.OrdinalIgnoreCase);
        clonedTab.ColumnWidths ??= new Dictionary<int, float>();
        clonedTab.RowHeights ??= new Dictionary<int, float>();
        clonedTab.Name = this.BuildUniqueTabName(document, string.IsNullOrWhiteSpace(preset.Name) ? clonedTab.Name : preset.Name);

        this.PushUndoSnapshot();
        document.Tabs.Add(clonedTab);
        document.ActiveTabIndex = document.Tabs.Count - 1;
        this.pendingTabDeletePopupPositions.Clear();
        this.ClearSelection();
        this.SelectSingleCell("R1C1");
        this.statusMessage = $"Added preset tab {clonedTab.Name}.";
    }

    private void SaveActiveTabAsPreset(string presetName)
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var activeTab = this.GetActiveTab();
        if (activeTab is null)
        {
            return;
        }

        var trimmedName = presetName.Trim();
        if (string.IsNullOrWhiteSpace(trimmedName))
        {
            return;
        }

        var settings = this.GetSheetSettings();
        settings.TabPresets ??= new List<SheetTabPreset>();
        var clonedTab = JsonSerializer.Deserialize<SheetTabData>(JsonSerializer.Serialize(activeTab)) ?? new SheetTabData();
        clonedTab.Name = trimmedName;
        clonedTab.Cells ??= new Dictionary<string, SheetCellData>(StringComparer.OrdinalIgnoreCase);
        clonedTab.CellFeatures ??= new Dictionary<string, CellFeatureBundle>(StringComparer.OrdinalIgnoreCase);
        clonedTab.ColumnWidths ??= new Dictionary<int, float>();
        clonedTab.RowHeights ??= new Dictionary<int, float>();

        var existing = settings.TabPresets.FirstOrDefault(x => string.Equals(x.Name, trimmedName, StringComparison.OrdinalIgnoreCase));
        if (existing is null)
        {
            settings.TabPresets.Add(new SheetTabPreset
            {
                Name = trimmedName,
                Tab = clonedTab,
            });
        }
        else
        {
            existing.Name = trimmedName;
            existing.Tab = clonedTab;
        }

        this.statusMessage = $"Saved preset {trimmedName}.";
    }

    private string BuildUniqueTabName(SheetDocument document, string baseName)
    {
        var trimmedBaseName = string.IsNullOrWhiteSpace(baseName) ? "Sheet" : baseName.Trim();
        if (!document.Tabs.Any(tab => string.Equals(tab.Name, trimmedBaseName, StringComparison.OrdinalIgnoreCase)))
        {
            return trimmedBaseName;
        }

        var suffix = 2;
        while (document.Tabs.Any(tab => string.Equals(tab.Name, $"{trimmedBaseName} {suffix}", StringComparison.OrdinalIgnoreCase)))
        {
            suffix++;
        }

        return $"{trimmedBaseName} {suffix}";
    }

    private static string BuildChatNotificationAnchor(List<SheetChatMessage>? messages)
    {
        var latest = messages?
            .OrderBy(x => x.TimestampUtc)
            .LastOrDefault();

        if (latest is null)
        {
            return string.Empty;
        }

        return string.IsNullOrWhiteSpace(latest.Id)
            ? $"{latest.AuthorUserId}|{latest.TimestampUtc.UtcDateTime.Ticks}|{latest.Message}"
            : latest.Id;
    }

    private void SyncChatNotificationAnchor(List<SheetChatMessage>? messages)
    {
        this.lastObservedChatNotificationAnchor = BuildChatNotificationAnchor(messages);
    }

    private bool IsMentionedInMessage(SheetChatMessage message)
    {
        var text = message.Message ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        if (text.Contains("@everyone", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var candidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var fullName = this.GetSafeCharacterFullName();
        if (!string.IsNullOrWhiteSpace(fullName))
        {
            candidates.Add(fullName);
            var firstName = fullName.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(firstName))
            {
                candidates.Add(firstName);
            }
        }

        var discordName = this.GetSafeDiscordDisplayName();
        if (!string.IsNullOrWhiteSpace(discordName))
        {
            candidates.Add(discordName);
        }

        return candidates.Any(candidate => text.Contains($"@{candidate}", StringComparison.OrdinalIgnoreCase));
    }

    private void HandleChatNotification(List<SheetChatMessage>? messages)
    {
        var latest = messages?
            .OrderBy(x => x.TimestampUtc)
            .LastOrDefault();
        var newAnchor = BuildChatNotificationAnchor(messages);

        if (string.IsNullOrWhiteSpace(this.lastObservedChatNotificationAnchor))
        {
            this.lastObservedChatNotificationAnchor = newAnchor;
            return;
        }

        if (string.IsNullOrWhiteSpace(newAnchor) || string.Equals(newAnchor, this.lastObservedChatNotificationAnchor, StringComparison.Ordinal))
        {
            return;
        }

        this.lastObservedChatNotificationAnchor = newAnchor;
        if (latest is null || string.Equals(latest.AuthorUserId, this.configuration.UserId, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var mention = this.IsMentionedInMessage(latest);
        if (!this.chatInputActive)
        {
            this.unreadChatCount++;
            this.chatHasUnreadAttention = true;
        }

        this.sidebarChatScrollToBottomPending = true;
        this.popupChatScrollToBottomPending = true;
        this.TryPlayChatNotificationSound(mention);
        if (mention)
        {
            this.TryShowMentionToast(latest);
        }
    }

    private void TryPlayChatNotificationSound(bool mention = false)
    {
        try
        {
            Task.Run(() =>
            {
                try
                {
                    PlaySound(mention ? "SystemExclamation" : "SystemAsterisk", 0, SoundAlias | SoundAsync | SoundNoDefault);
                }
                catch
                {
                }

                try
                {
                    MessageBeep(mention ? 0x00000030u : 0x00000040u);
                }
                catch
                {
                }

                try
                {
                    if (mention)
                    {
                        SystemSounds.Exclamation.Play();
                    }
                    else
                    {
                        SystemSounds.Asterisk.Play();
                    }
                }
                catch
                {
                }

                try
                {
                    Console.Beep(mention ? 1150 : 900, mention ? 180 : 120);
                }
                catch
                {
                }
            });
        }
        catch
        {
        }
    }

    private void TryShowMentionToast(SheetChatMessage message)
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var anchor = $"{message.Id}|{message.TimestampUtc.ToUnixTimeMilliseconds()}|{message.AuthorUserId}";
        if (string.Equals(anchor, this.lastMentionToastAnchor, StringComparison.Ordinal))
        {
            return;
        }

        this.lastMentionToastAnchor = anchor;
        var content = $"[{DateTime.Now:HH:mm}] {message.AuthorName} Mentioned you in the sheet {this.currentSheet.Title}";
        try
        {
            Plugin.NotificationManager.AddNotification(new Notification
            {
                Title = "EZSheets",
                Content = content,
                Type = NotificationType.None,
                InitialDuration = TimeSpan.FromSeconds(5),
            });
        }
        catch
        {
        }

        try
        {
            Plugin.ToastGui.ShowNormal(content);
        }
        catch
        {
        }
    }

    private void DrawMemberRemovedPopup()
    {
        if (!this.memberRemovedPopupOpen)
        {
            return;
        }

        var open = this.memberRemovedPopupOpen;
        ImGui.SetNextWindowSize(new Vector2(360f, 132f), ImGuiCond.Appearing);
        if (ImGui.Begin("User removed###EZSheetsMemberRemoved", ref open, ImGuiWindowFlags.NoCollapse | ImGuiWindowFlags.NoResize))
        {
            ImGui.TextWrapped($"You removed {this.memberRemovedName} access for this sheet.");
            ImGui.TextWrapped($"{this.memberRemovedName} is now on the Blocklist");
            ImGui.Spacing();
            if (this.DrawStyledButton("member-removed-ok", "OK", AccentBlue, WhiteText, new Vector2(72f, 24f)))
            {
                this.memberRemovedPopupOpen = false;
            }
        }

        ImGui.End();
        this.memberRemovedPopupOpen = open;
    }

    private bool AreChatMessagesEquivalent(List<SheetChatMessage>? lhs, List<SheetChatMessage>? rhs)
    {
        var left = lhs ?? new List<SheetChatMessage>();
        var right = rhs ?? new List<SheetChatMessage>();
        if (left.Count != right.Count)
        {
            return false;
        }

        for (var index = 0; index < left.Count; index++)
        {
            var a = left[index];
            var b = right[index];
            if (!string.Equals(a.Id, b.Id, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(a.AuthorUserId, b.AuthorUserId, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(a.AuthorName, b.AuthorName, StringComparison.Ordinal)
                || !string.Equals(a.Message, b.Message, StringComparison.Ordinal)
                || a.TimestampUtc != b.TimestampUtc)
            {
                return false;
            }
        }

        return true;
    }

    private void DrawFeatureBorders(ImDrawListPtr drawList, Vector2 min, Vector2 max, SheetCellData cell)
    {
        if (cell.Borders == CellBorderFlags.None)
        {
            return;
        }

        var color = cell.BorderColor != 0
            ? cell.BorderColor
            : ImGui.ColorConvertFloat4ToU32(new Vector4(1f, 1f, 1f, 0.90f));
        color = (color & 0x00FFFFFFu) | 0xE6000000u;
        const float thickness = 3.3f;
        var half = thickness * 0.5f;
        if (cell.Borders.HasFlag(CellBorderFlags.Left))
        {
            drawList.AddLine(new Vector2(min.X + half, min.Y), new Vector2(min.X + half, max.Y), color, thickness);
        }

        if (cell.Borders.HasFlag(CellBorderFlags.Top))
        {
            drawList.AddLine(new Vector2(min.X, min.Y + half), new Vector2(max.X, min.Y + half), color, thickness);
        }

        if (cell.Borders.HasFlag(CellBorderFlags.Right))
        {
            drawList.AddLine(new Vector2(max.X - half, min.Y), new Vector2(max.X - half, max.Y), color, thickness);
        }

        if (cell.Borders.HasFlag(CellBorderFlags.Bottom))
        {
            drawList.AddLine(new Vector2(min.X, max.Y - half), new Vector2(max.X, max.Y - half), color, thickness);
        }
    }

    private readonly record struct SearchResultItem(int TabIndex, string TabName, string CellKey, string CellLabel, string Preview);
}
