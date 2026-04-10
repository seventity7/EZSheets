using System.Diagnostics;
using System.Globalization;
using System.Numerics;
using System.Text.Json;
using ClosedXML.Excel;
using Dalamud.Bindings.ImGui;
using WinForms = System.Windows.Forms;
using Dalamud.Interface.Windowing;
using EZSheets.Models;
using EZSheets.Services;

namespace EZSheets.Windows;

public sealed partial class MainWindow : Window, IDisposable
{
    private static readonly Vector4 AccentOrange = HexToVector4("#ff9100");
    private static readonly Vector4 AccentBlue = HexToVector4("#0074e0");
    private static readonly Vector4 AccentGreen = HexToVector4("#83db18");
    private static readonly Vector4 AccentTabAdd = HexToVector4("#49a800");
    private static readonly Vector4 AccentSignOut = HexToVector4("#c2311b");
    private static readonly Vector4 AccentCreate = HexToVector4("#41a308");
    private static readonly Vector4 AccentJoin = HexToVector4("#0872a3");
    private static readonly Vector4 AccentSaveLocal = HexToVector4("#94b812");
    private static readonly Vector4 AccentDownload = HexToVector4("#12b899");
    private static readonly Vector4 AccentSaveCloud = HexToVector4("#1291b8");
    private static readonly Vector4 AccentDelete = HexToVector4("#ba0606");
    private static readonly Vector4 AccentCopy = HexToVector4("#46ad1a");
    private static readonly Vector4 AccentArrowOpen = HexToVector4("#db4218");
    private static readonly Vector4 AccentArrowClosed = HexToVector4("#db8318");
    private static readonly Vector4 GrayDisabled = HexToVector4("#8a8a8a");
    private static readonly Vector4 WhiteText = new(1f, 1f, 1f, 1f);

    private readonly Configuration configuration;
    private readonly SupabaseRestClient supabase;
    private readonly SemaphoreSlim sheetSaveGate = new(1, 1);

    private readonly List<SheetSummary> accessibleSheets = new();
    private readonly object accessibleSheetsSync = new();
    private readonly HashSet<string> selectedCellKeys = new(StringComparer.OrdinalIgnoreCase);

    private string joinCode = string.Empty;
    private string newSheetTitle = "New Sheet";
    private int newRows = 30;
    private int newCols = 12;
    private bool newSheetDefaultEditor;

    private RemoteSheet? currentSheet;
    private SheetAccessRole currentRole = SheetAccessRole.Viewer;
    private string selectedCellKey = "R1C1";
    private string statusMessage = "Ready.";
    private int actionInProgressCount;
    private bool actionInProgress => Volatile.Read(ref this.actionInProgressCount) > 0;
    private bool localSavePopupOpen;
    private bool localSavePopupRequested;
    private string lastSavedLocalPath = string.Empty;
    private string? editingCellKey;
    private string editingCellBuffer = string.Empty;
    private bool editingFocusRequested;
    private int editingStartedFrame;
    private bool helpWindowOpen;
    private readonly Stack<string> undoHistory = new();
    private readonly List<MergedTextOverlay> pendingMergedTextOverlays = new();
    private readonly List<MergedSelectionOverlay> pendingMergedSelectionOverlays = new();
    private MergedEditOverlay? pendingMergedEditOverlay;
    private readonly object importDialogSync = new();
    private bool importDialogOpen;
    private string? pendingImportPath;
    private string? pendingImportError;
    private bool pendingClearCurrentSheet;
    private string? pendingClearedSheetId;
    private Vector2 gridClipMin;
    private Vector2 gridClipMax;
    private DateTimeOffset copyCodeFeedbackUntilUtc = DateTimeOffset.MinValue;
    private DateTimeOffset copyUniqueFeedbackUntilUtc = DateTimeOffset.MinValue;
    private SheetUniqueCodeInfo? currentOwnerUniqueCode;
    private bool hideSharedCode;
    private bool discordLoginInProgress;
    private string? uniqueCodeEnsureSheetId;
    private bool uniqueCodeEnsureInFlight;
    private bool restoreWindowLayoutPending = true;
    private Vector2 lastPersistedWindowPos = new(float.NaN, float.NaN);
    private Vector2 lastPersistedWindowSize = new(float.NaN, float.NaN);
    private DateTimeOffset lastWindowLayoutPersistUtc = DateTimeOffset.MinValue;
    private bool sheetSearchPanelOpen;
    private string sheetSearchQuery = string.Empty;
    private bool signOutConfirmPopupOpen;
    private Vector2 signOutConfirmAnchor = Vector2.Zero;

    public MainWindow(Configuration configuration, SupabaseRestClient supabase, Action openConfigAction)
        : base("EZSheets###EZSheetsMainWindow")
    {
        _ = openConfigAction;
        this.SizeCondition = ImGuiCond.FirstUseEver;
        this.Size = new Vector2(1030, 620);
        this.configuration = configuration;
        if (this.configuration.HasSavedMainWindowLayout
            && this.configuration.MainWindowWidth >= 640f
            && this.configuration.MainWindowHeight >= 420f)
        {
            this.Size = new Vector2(this.configuration.MainWindowWidth, this.configuration.MainWindowHeight);
        }
        this.supabase = supabase;
        this.selectedCellKeys.Add(this.selectedCellKey);
        this.InitializeAdvancedFeatures();
    }

    public void Dispose()
    {
        this.sheetSaveGate.Dispose();
    }

    private sealed record MergedTextOverlay(ImDrawListPtr DrawList, Vector2 Min, Vector2 Max, SheetCellData Cell, string Text);
    private sealed record MergedSelectionOverlay(ImDrawListPtr DrawList, Vector2 Min, Vector2 Max);
    private sealed record MergedEditOverlay(Vector2 Min, Vector2 Max, string CellKey);

    public async Task TryRestoreSessionAndLoadAsync()
    {
        if (!this.supabase.HasConfiguredProject)
        {
            return;
        }

        if (await this.supabase.RestoreSessionAsync().ConfigureAwait(false))
        {
            var displayName = this.GetSafeCharacterFullName();
            this.statusMessage = $"Restored session for {displayName}.";
            await this.ReloadSheetListAsync().ConfigureAwait(false);

            if (!string.IsNullOrWhiteSpace(this.configuration.LastOpenedSheetId))
            {
                try
                {
                    await this.LoadSheetAsync(this.configuration.LastOpenedSheetId).ConfigureAwait(false);
                }
                catch
                {
                    this.statusMessage = "Session restored, but the previous sheet could not be loaded automatically.";
                }
            }
        }
    }

    public override void Draw()
    {
        this.TryRestoreWindowLayout();
        this.ProcessPendingUiActions();
        this.UpdateAdvancedRuntime();
        this.DrawStatusBar();

        var avail = ImGui.GetContentRegionAvail();
        var leftWidth = this.configuration.SidebarCollapsed ? 34f : MathF.Min(285f, MathF.Max(245f, avail.X * 0.23f));

        if (ImGui.BeginChild("##leftPane", new Vector2(leftWidth, 0), true))
        {
            this.DrawSidebar();
        }

        ImGui.EndChild();
        ImGui.SameLine();

        if (ImGui.BeginChild("##rightPane", new Vector2(0, 0), true))
        {
            this.HandleEditingFocusLoss();
            this.DrawSheetEditor();
        }

        ImGui.EndChild();

        this.DrawDeletePopup();
        this.DrawLocalSavePopup();
        this.DrawHelpWindow();
        this.DrawAdvancedWindows();
        this.DrawSidebarSignOutConfirmPopup();
        this.TryPersistWindowLayout();
    }

    private void TryRestoreWindowLayout()
    {
        if (!this.restoreWindowLayoutPending)
        {
            return;
        }

        this.restoreWindowLayoutPending = false;
        if (!this.configuration.HasSavedMainWindowLayout)
        {
            return;
        }

        var targetSize = new Vector2(
            MathF.Max(640f, this.configuration.MainWindowWidth),
            MathF.Max(420f, this.configuration.MainWindowHeight));
        var targetPos = new Vector2(this.configuration.MainWindowPosX, this.configuration.MainWindowPosY);
        ImGui.SetWindowSize(targetSize, ImGuiCond.Always);
        ImGui.SetWindowPos(targetPos, ImGuiCond.Always);
        this.lastPersistedWindowPos = targetPos;
        this.lastPersistedWindowSize = targetSize;
    }

    private void TryPersistWindowLayout()
    {
        var windowPos = ImGui.GetWindowPos();
        var windowSize = ImGui.GetWindowSize();
        if (windowSize.X < 1f || windowSize.Y < 1f)
        {
            return;
        }

        var posChanged = float.IsNaN(this.lastPersistedWindowPos.X)
            || Vector2.DistanceSquared(windowPos, this.lastPersistedWindowPos) > 0.25f;
        var sizeChanged = float.IsNaN(this.lastPersistedWindowSize.X)
            || Vector2.DistanceSquared(windowSize, this.lastPersistedWindowSize) > 0.25f;
        if (!posChanged && !sizeChanged)
        {
            return;
        }

        var now = DateTimeOffset.UtcNow;
        if (now - this.lastWindowLayoutPersistUtc < TimeSpan.FromMilliseconds(250))
        {
            return;
        }

        this.lastWindowLayoutPersistUtc = now;
        this.lastPersistedWindowPos = windowPos;
        this.lastPersistedWindowSize = windowSize;
        this.configuration.MainWindowPosX = windowPos.X;
        this.configuration.MainWindowPosY = windowPos.Y;
        this.configuration.MainWindowWidth = windowSize.X;
        this.configuration.MainWindowHeight = windowSize.Y;
        this.configuration.HasSavedMainWindowLayout = true;
        this.configuration.Save();
    }

    private void ProcessPendingUiActions()
    {
        if (!this.pendingClearCurrentSheet)
        {
            return;
        }

        this.pendingClearCurrentSheet = false;
        var removedId = this.pendingClearedSheetId ?? string.Empty;
        this.pendingClearedSheetId = null;

        if (this.currentSheet is not null && string.Equals(this.currentSheet.Id, removedId, StringComparison.OrdinalIgnoreCase))
        {
            this.currentSheet = null;
            this.currentRole = SheetAccessRole.Viewer;
            this.ClearSelection();
        }

        if (!string.IsNullOrWhiteSpace(removedId))
        {
            this.accessibleSheets.RemoveAll(s => string.Equals(s.Id, removedId, StringComparison.OrdinalIgnoreCase));
        }
    }

    private void HandleEditingFocusLoss()
    {
        if (string.IsNullOrWhiteSpace(this.editingCellKey))
        {
            return;
        }

        if (ImGui.IsWindowFocused(ImGuiFocusedFlags.RootAndChildWindows))
        {
            return;
        }

        this.CommitAnyInCellEdit();
    }

    private void DrawStatusBar()
    {
        ImGui.TextColored(this.actionInProgress ? HexToVector4("#ffd54f") : HexToVector4("#79d279"), this.actionInProgress ? "Working..." : "Idle");
        ImGui.SameLine();
        ImGui.TextWrapped(this.statusMessage);
        ImGui.Spacing();
    }

    private void DrawSidebar()
    {
        if (this.configuration.SidebarCollapsed)
        {
            if (this.DrawStyledButton("sidebar-open", "→", AccentArrowClosed, WhiteText, new Vector2(32f, 28f)))
            {
                this.configuration.SidebarCollapsed = false;
                this.configuration.Save();
            }

            return;
        }

        this.DrawSidebarHeader();
        this.DrawSectionSpacer(4f);
        this.DrawAuthSection();
        var showHomeSections = this.currentSheet is null || this.homeViewActive;
        if (showHomeSections)
        {
            this.DrawCompactSectionDivider();
            this.DrawCreateAndJoinSection();
        }
        if (this.supabase.IsAuthenticated && this.sheetListPinned)
        {
            this.DrawCompactSectionDivider();
            this.DrawSheetList(false, showHomeSections);
        }

        if (this.currentSheet is not null && !this.homeViewActive)
        {
            this.DrawSidebarPresencePanel();
            this.DrawSidebarActivityFeedPanel();
            this.DrawSidebarChatPanel();
        }
    }

    private void DrawSidebarHeader()
    {
        var start = ImGui.GetCursorPos();
        var availableWidth = ImGui.GetContentRegionAvail().X;
        var iconSize = new Vector2(24f, 24f);
        var leftX = start.X;
        var closeX = start.X + MathF.Max(0f, availableWidth - iconSize.X);
        var homeX = closeX - iconSize.X - 6f;
        var titleX = leftX;

        if (this.supabase.IsAuthenticated)
        {
            ImGui.SetCursorPos(new Vector2(leftX, start.Y - 1f));
            if (this.DrawIconButton("sidebar-signout-icon", iconSize, AccentSignOut, WhiteText, DrawPowerGlyph))
            {
                this.signOutConfirmAnchor = new Vector2(ImGui.GetItemRectMin().X, ImGui.GetItemRectMin().Y);
                this.signOutConfirmPopupOpen = true;
            }
            if (ImGui.IsItemHovered())
            {
                ImGui.SetTooltip("Sign out");
            }
            titleX += iconSize.X + 6f;
        }

        ImGui.SetCursorPos(new Vector2(titleX, start.Y));
        this.DrawFakeBoldText("Account", AccentOrange);

        if (this.currentSheet is not null)
        {
            ImGui.SetCursorPos(new Vector2(homeX, start.Y - 1f));
            if (this.DrawIconButton("sidebar-home-toggle", iconSize, HexToVector4("#8f3859"), WhiteText, this.homeViewActive ? DrawOverviewGlyph : DrawHomeGlyph))
            {
                this.homeViewActive = !this.homeViewActive;
            }
            if (ImGui.IsItemHovered())
            {
                ImGui.SetTooltip(this.homeViewActive ? "Return to opened sheet" : "Go to home screen");
            }
        }

        ImGui.SetCursorPos(new Vector2(closeX, start.Y - 1f));
        if (this.DrawStyledButton("sidebar-close", "←", AccentArrowOpen, WhiteText, iconSize))
        {
            this.configuration.SidebarCollapsed = true;
            this.configuration.Save();
        }
    }


    private void DrawSidebarSignOutConfirmPopup()
    {
        if (!this.signOutConfirmPopupOpen || !this.supabase.IsAuthenticated)
        {
            return;
        }

        var popupSize = new Vector2(270f, 96f);
        var popupPos = new Vector2(this.signOutConfirmAnchor.X, this.signOutConfirmAnchor.Y - popupSize.Y - 8f);
        var popupMax = popupPos + popupSize;
        ImGui.SetNextWindowPos(popupPos, ImGuiCond.Always);
        ImGui.SetNextWindowSize(popupSize, ImGuiCond.Always);
        var open = this.signOutConfirmPopupOpen;
        if (!ImGui.Begin("##sidebar-signout-confirm", ref open, ImGuiWindowFlags.NoDecoration | ImGuiWindowFlags.NoResize | ImGuiWindowFlags.NoSavedSettings | ImGuiWindowFlags.NoMove))
        {
            ImGui.End();
            this.signOutConfirmPopupOpen = open;
            return;
        }

        if ((ImGui.IsMouseClicked(ImGuiMouseButton.Left) || ImGui.IsMouseClicked(ImGuiMouseButton.Right))
            && !ImGui.IsMouseHoveringRect(popupPos, popupMax, true))
        {
            ImGui.End();
            this.signOutConfirmPopupOpen = false;
            return;
        }

        var displayName = this.GetSafeCharacterFullName();
        ImGui.TextWrapped("Make sure to save your sheet to Cloud.");
        ImGui.TextWrapped($"Log Off as {displayName}?");
        ImGui.Spacing();
        if (this.DrawStyledButton("sidebar-signout-confirm-yes", "Yes", AccentSignOut, WhiteText, new Vector2(58f, 22f)))
        {
            ImGui.End();
            this.signOutConfirmPopupOpen = false;
            _ = this.RunActionAsync(async () =>
            {
                await this.supabase.SignOutAsync().ConfigureAwait(false);
                this.accessibleSheets.Clear();
                this.currentSheet = null;
                this.currentRole = SheetAccessRole.Viewer;
                this.sheetListWindowOpen = false;
                this.sheetListPinned = true;
                this.presenceWindowOpen = false;
                this.presencePinned = true;
                this.homeViewActive = false;
                this.ClearSelection();
                this.statusMessage = "Signed out.";
            }, "Sign out failed.");
            return;
        }

        ImGui.SameLine(0f, 6f);
        if (this.DrawStyledButton("sidebar-signout-confirm-no", "No", HexToVector4("#666666"), WhiteText, new Vector2(52f, 22f)))
        {
            ImGui.End();
            this.signOutConfirmPopupOpen = false;
            return;
        }

        ImGui.End();
        this.signOutConfirmPopupOpen = open;
    }

    private void DrawAuthSection()
    {
        if (this.supabase.IsAuthenticated)
        {
            return;
        }

        ImGui.PushStyleColor(ImGuiCol.Text, AccentOrange);
        ImGui.TextWrapped("A browser window will open. Finish the Discord login there, then return to the game.");
        ImGui.PopStyleColor();
        ImGui.Spacing();

        var signInColor = this.discordLoginInProgress ? GrayDisabled : AccentBlue;
        var signInLabel = this.discordLoginInProgress ? "Waiting for Discord..." : "Sign In with Discord";
        if (this.DrawStyledButton("discord-login", signInLabel, signInColor, WhiteText, new Vector2(-1, 28)) && !this.discordLoginInProgress)
        {
            this.discordLoginInProgress = true;
            _ = this.RunActionAsync(async () =>
            {
                try
                {
                    _ = await this.supabase.SignInWithDiscordAsync().ConfigureAwait(false);
                    var displayName = this.GetSafeCharacterFullName();
                    this.statusMessage = $"Signed in as {displayName}. Loading sheets...";
                    _ = Task.Run(async () =>
                    {
                        try
                        {
                            await this.ReloadSheetListAsync().ConfigureAwait(false);
                            this.statusMessage = $"Signed in as {displayName}.";
                        }
                        catch (Exception ex)
                        {
                            this.statusMessage = string.IsNullOrWhiteSpace(ex.Message) ? "Signed in, but could not load sheet list yet." : ex.Message;
                        }
                    });
                }
                finally
                {
                    this.discordLoginInProgress = false;
                }
            }, "Discord sign-in failed.");
        }

        if (this.discordLoginInProgress || this.supabase.IsDiscordLoginInProgress)
        {
            ImGui.Spacing();
            if (this.DrawStyledButton("discord-login-cancel", "Cancel", AccentDelete, WhiteText, new Vector2(-1, 26f)))
            {
                this.supabase.CancelActiveDiscordLogin();
                this.discordLoginInProgress = false;
                this.statusMessage = "Discord login cancelled.";
            }
        }

        ImGui.Spacing();
        ImGui.PushStyleColor(ImGuiCol.Text, GrayDisabled);
        ImGui.TextWrapped("If the browser does not return automatically, wait a moment and then come back to the game window.");
        ImGui.PopStyleColor();
    }

    private void DrawCreateAndJoinSection()
    {
        this.DrawFakeBoldText("EZ Sheets", AccentOrange);
        ImGui.Spacing();

        if (!this.supabase.IsAuthenticated)
        {
            ImGui.PushStyleColor(ImGuiCol.Text, WhiteText);
            ImGui.TextWrapped("Login to be able to create/view sheets");
            ImGui.PopStyleColor();
            return;
        }

        if (this.currentSheet is null || this.homeViewActive)
        {
            ImGui.TextUnformatted("Title");
            ImGui.SetNextItemWidth(-1);
            ImGui.InputText("##NewSheetTitle", ref this.newSheetTitle, 128);
            this.DrawTemplatePicker();
            ImGui.SliderInt("Rows", ref this.newRows, 5, 100);
            ImGui.SliderInt("Columns", ref this.newCols, 3, 26);
            ImGui.Checkbox("Share code grants Editor access", ref this.newSheetDefaultEditor);

            if (this.DrawStyledButton("create-sheet", "Create New Sheet", AccentCreate, WhiteText, new Vector2(-1, 24)))
            {
                _ = this.RunActionAsync(async () =>
                {
                    var created = await this.supabase.CreateSheetAsync(
                            string.IsNullOrWhiteSpace(this.newSheetTitle) ? "New Sheet" : this.newSheetTitle.Trim(),
                            this.newRows,
                            this.newCols,
                            this.newSheetDefaultEditor ? "editor" : "viewer",
                            this.CreateDocumentFromSelectedTemplate())
                        .ConfigureAwait(false);

                    await this.ReloadSheetListAsync().ConfigureAwait(false);
                    await this.LoadSheetAsync(created.Id).ConfigureAwait(false);
                    this.statusMessage = $"Created new sheet {created.Title}.";
                }, "Could not create sheet.");
            }
        }
        else
        {
            if (this.DrawStyledButton("create-sheet-inline", "Create New Sheet", AccentCreate, WhiteText, new Vector2(-1, 24)))
            {
                ImGui.OpenPopup("Create New Sheet Popup");
            }

            if (ImGui.BeginPopup("Create New Sheet Popup"))
            {
                ImGui.TextUnformatted("Title");
                ImGui.SetNextItemWidth(230f);
                ImGui.InputText("##NewSheetTitlePopup", ref this.newSheetTitle, 128);
                this.DrawTemplatePicker();
                ImGui.SliderInt("Rows", ref this.newRows, 5, 100);
                ImGui.SliderInt("Columns", ref this.newCols, 3, 26);
                ImGui.Checkbox("Share code grants Editor access", ref this.newSheetDefaultEditor);
                if (this.DrawStyledButton("create-sheet-popup-confirm", "Create", AccentCreate, WhiteText, new Vector2(120f, 24f)))
                {
                    ImGui.CloseCurrentPopup();
                    _ = this.RunActionAsync(async () =>
                    {
                        var created = await this.supabase.CreateSheetAsync(
                                string.IsNullOrWhiteSpace(this.newSheetTitle) ? "New Sheet" : this.newSheetTitle.Trim(),
                                this.newRows,
                                this.newCols,
                                this.newSheetDefaultEditor ? "editor" : "viewer",
                                this.CreateDocumentFromSelectedTemplate())
                            .ConfigureAwait(false);

                        await this.ReloadSheetListAsync().ConfigureAwait(false);
                        await this.LoadSheetAsync(created.Id).ConfigureAwait(false);
                        this.statusMessage = $"Created new sheet {created.Title}.";
                    }, "Could not create sheet.");
                }
                ImGui.EndPopup();
            }
        }

        this.DrawSectionDivider();
        ImGui.PushStyleColor(ImGuiCol.Text, AccentOrange);
        ImGui.TextUnformatted("Shared Code");
        ImGui.PopStyleColor();
        ImGui.SetNextItemWidth(-1);
        ImGui.InputText("##JoinCode", ref this.joinCode, 64);
        if (string.IsNullOrWhiteSpace(this.joinCode) && !ImGui.IsItemActive())
        {
            var min = ImGui.GetItemRectMin();
            var drawList = ImGui.GetWindowDrawList();
            drawList.AddText(new Vector2(min.X + 7f, min.Y + 4f), ImGui.ColorConvertFloat4ToU32(GrayDisabled), "Example: 86SSFCG2P7");
        }

        if (this.DrawStyledButton("join-by-code", "Access shared sheet", AccentJoin, WhiteText, new Vector2(-1, 24)))
        {
            var code = this.joinCode.Trim();
            _ = this.RunActionAsync(async () =>
            {
                try
                {
                    var joinedId = await this.supabase.JoinSheetByCodeAsync(code).ConfigureAwait(false);
                    await this.ReloadSheetListAsync().ConfigureAwait(false);
                    await this.LoadSheetAsync(joinedId).ConfigureAwait(false);
                    this.statusMessage = $"Accessed shared sheet {code.ToUpperInvariant()}.";
                    this.joinCode = string.Empty;
                }
                catch (SupabaseApiException ex) when (ex.Message.Contains("blocked", StringComparison.OrdinalIgnoreCase))
                {
                    this.blockedJoinPopupOpen = true;
                    this.blockedJoinMessage = "You are blocked from this sheet.";
                    this.statusMessage = this.blockedJoinMessage;
                }
            }, "Could not access that shared sheet.");
        }
    }


    private void DrawSheetList(bool detached, bool expandToAvailable = false)
    {
        if (!this.supabase.IsAuthenticated)
        {
            return;
        }

        var start = ImGui.GetCursorPos();
        this.DrawFakeBoldText("Available Sheets", AccentOrange);
        var btnSize = new Vector2(20f, 20f);
        var rightX = start.X + MathF.Max(0f, ImGui.GetContentRegionAvail().X - btnSize.X);
        ImGui.SetCursorPos(new Vector2(rightX, start.Y - 1f));
        if (this.DrawStyledButton(detached ? "sheet-list-pin-window" : "sheet-list-pop-pin", "♯", HexToVector4("#3b3b3b"), WhiteText, btnSize))
        {
            this.sheetListPinned = detached;
            this.sheetListWindowOpen = !detached;
        }
        if (ImGui.IsItemHovered())
        {
            ImGui.SetTooltip(detached ? "Pin available sheets" : "Pop out available sheets");
        }

        ImGui.Spacing();

        if (this.DrawStyledButton(detached ? "refresh-list-window" : "refresh-list", "Refresh List", HexToVector4("#5c15ad"), WhiteText, new Vector2(-1, 26)) && this.supabase.IsAuthenticated)
        {
            _ = this.RunActionAsync(this.ReloadSheetListAsync, "Could not refresh the sheet list.");
        }

        if (this.accessibleSheets.Count >= 3 && this.DrawStyledButton(detached ? "search-sheets-toggle-window" : "search-sheets-toggle", "Search Sheets", HexToVector4("#5a3f9e"), WhiteText, new Vector2(-1f, 24f)))
        {
            this.sheetSearchPanelOpen = !this.sheetSearchPanelOpen;
            if (!this.sheetSearchPanelOpen)
            {
                this.sheetSearchQuery = string.Empty;
            }
        }

        if (this.sheetSearchPanelOpen)
        {
            if (ImGui.BeginChild(detached ? "##sheet-search-panel-window" : "##sheet-search-panel", new Vector2(-1f, 64f), true, ImGuiWindowFlags.NoScrollbar | ImGuiWindowFlags.NoScrollWithMouse))
            {
                ImGui.TextDisabled("Search by sheet name, category, role, owner or code");
                ImGui.SetNextItemWidth(-1f);
                ImGui.InputText(detached ? "##sheet-search-query-window" : "##sheet-search-query", ref this.sheetSearchQuery, 128);
            }
            ImGui.EndChild();
        }

        ImGui.Spacing();

        if (this.accessibleSheets.Count < 3)
        {
            this.sheetSearchPanelOpen = false;
            this.sheetSearchQuery = string.Empty;
        }

        List<SheetSummary> sheetSnapshot;
        lock (this.accessibleSheetsSync)
        {
            sheetSnapshot = this.accessibleSheets
                .OrderByDescending(sheet => this.IsSheetFavorited(sheet.Id))
                .ThenByDescending(sheet => this.configuration.RecentSheetAccessTicks.TryGetValue(sheet.Id, out var recentTicks) ? recentTicks : 0L)
                .ThenByDescending(sheet => sheet.UpdatedAt)
                .ThenByDescending(sheet => sheet.CreatedAt)
                .ThenByDescending(sheet => sheet.Id, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        if (!string.IsNullOrWhiteSpace(this.sheetSearchQuery))
        {
            sheetSnapshot = sheetSnapshot
                .Where(this.MatchesSheetSearchQuery)
                .ToList();
        }

        if (sheetSnapshot.Count == 0)
        {
            ImGui.PushStyleColor(ImGuiCol.Text, GrayDisabled);
            ImGui.TextWrapped(string.IsNullOrWhiteSpace(this.sheetSearchQuery) ? "No sheets available yet." : "No sheets matched your search.");
            ImGui.PopStyleColor();
            return;
        }

        var estimatedSheetItemHeight = 92f;
        var maxVisibleSheetHeight = detached || expandToAvailable ? MathF.Max(220f, ImGui.GetContentRegionAvail().Y - 4f) : estimatedSheetItemHeight * 2f;
        if (ImGui.BeginChild(detached ? "##SheetListScrollWindow" : "##SheetListScroll", new Vector2(0f, maxVisibleSheetHeight), false, ImGuiWindowFlags.AlwaysVerticalScrollbar))
        {
            var autoScrolledToSearchResult = false;
            foreach (var sheet in sheetSnapshot)
            {
                ImGui.PushID(sheet.Id);
                var color = this.GetSheetListColor(sheet.Id);
                var colorVector = ImGui.ColorConvertU32ToFloat4(color);

                if (this.DrawColorSwatchButton("sheet-color-swatch", colorVector, new Vector2(14f, 14f)))
                {
                    ImGui.OpenPopup("sheet-color-popup");
                }

                if (ImGui.BeginPopup("sheet-color-popup"))
                {
                    var pickerColor = colorVector;
                    if (ImGui.ColorPicker4("##picker", ref pickerColor, ImGuiColorEditFlags.NoSidePreview | ImGuiColorEditFlags.NoSmallPreview | ImGuiColorEditFlags.AlphaBar))
                    {
                        this.configuration.SheetListColors[sheet.Id] = ImGui.ColorConvertFloat4ToU32(pickerColor);
                        this.configuration.Save();
                        colorVector = pickerColor;
                    }

                    ImGui.EndPopup();
                }

                ImGui.SameLine();
                ImGui.PushStyleColor(ImGuiCol.Text, colorVector);
                var isSelected = this.currentSheet?.Id == sheet.Id;
                var cachedFavorite = this.IsSheetFavorited(sheet.Id);
                var cachedCategory = this.configuration.LocalSheets.TryGetValue(sheet.Id, out var cachedSheet) ? cachedSheet.Data.Settings.Category : string.Empty;
                var displayTitle = cachedFavorite ? "★ " + StripAutoVersionSuffix(sheet.Title) : StripAutoVersionSuffix(sheet.Title);
                var searchHighlight = !string.IsNullOrWhiteSpace(this.sheetSearchQuery) && this.MatchesSheetSearchQuery(sheet);
                if (searchHighlight)
                {
                    ImGui.PushStyleColor(ImGuiCol.Header, new Vector4(0.42f, 0.31f, 0.12f, 0.62f));
                    ImGui.PushStyleColor(ImGuiCol.HeaderHovered, new Vector4(0.58f, 0.42f, 0.12f, 0.72f));
                }
                if (ImGui.Selectable(displayTitle, isSelected))
                {
                    _ = this.RunActionAsync(async () => await this.LoadSheetAsync(sheet.Id).ConfigureAwait(false), "Could not load that sheet.");
                }
                if (searchHighlight)
                {
                    if (!autoScrolledToSearchResult)
                    {
                        ImGui.SetScrollHereY(0.15f);
                        autoScrolledToSearchResult = true;
                    }
                    ImGui.PopStyleColor(2);
                }

                ImGui.PopStyleColor();
                ImGui.PushStyleColor(ImGuiCol.Text, GrayDisabled);
                if (string.Equals(sheet.OwnerId, this.configuration.UserId, StringComparison.OrdinalIgnoreCase))
                {
                    ImGui.TextWrapped($"Share code: {sheet.Code}");
                }
                if (!string.IsNullOrWhiteSpace(cachedCategory))
                {
                    ImGui.TextWrapped($"Category: {cachedCategory}");
                }
                ImGui.TextWrapped($"Updated: {sheet.UpdatedAt.LocalDateTime:g}");
                var listRole = !string.IsNullOrWhiteSpace(sheet.UserRole)
                    ? sheet.UserRole
                    : (this.currentSheet is not null && string.Equals(this.currentSheet.Id, sheet.Id, StringComparison.OrdinalIgnoreCase)
                        ? this.GetDisplayedRoleLabel()
                        : (string.Equals(sheet.OwnerId, this.configuration.UserId, StringComparison.OrdinalIgnoreCase) ? "Owner" : string.Empty));
                if (!string.IsNullOrWhiteSpace(listRole))
                {
                    ImGui.TextWrapped($"Role: {listRole}");
                }
                ImGui.PopStyleColor();
                this.DrawSectionSpacer(4f);
                ImGui.Separator();
                this.DrawSectionSpacer(4f);
                ImGui.PopID();
            }
        }
        ImGui.EndChild();
    }

    private void DrawDetachedSheetListWindow()
    {
        if (!this.sheetListWindowOpen || !this.supabase.IsAuthenticated)
        {
            return;
        }

        var open = this.sheetListWindowOpen;
        this.TryRestoreDetachedWindowLayout("sheetlist");
        ImGui.SetNextWindowSize(new Vector2(360f, 330f), ImGuiCond.FirstUseEver);
        if (!ImGui.Begin("Available Sheets###EZSheetsSheetsList", ref open))
        {
            ImGui.End();
            this.sheetListWindowOpen = open;
            if (!open)
            {
                this.sheetListPinned = true;
            }
            return;
        }

        this.sheetListWindowOpen = open;
        this.TryPersistDetachedWindowLayout("sheetlist");
        this.DrawSheetList(true);
        ImGui.End();
        if (!open)
        {
            this.sheetListPinned = true;
        }
    }

    private bool MatchesSheetSearchQuery(SheetSummary sheet)
    {
        var query = this.sheetSearchQuery?.Trim();
        if (string.IsNullOrWhiteSpace(query))
        {
            return true;
        }

        var cachedCategory = this.configuration.LocalSheets.TryGetValue(sheet.Id, out var cachedSheet)
            ? cachedSheet.Data.Settings.Category ?? string.Empty
            : string.Empty;
        var ownerDisplay = string.Equals(sheet.OwnerId, this.configuration.UserId, StringComparison.OrdinalIgnoreCase)
            ? "owner"
            : sheet.OwnerId ?? string.Empty;
        var role = sheet.UserRole ?? string.Empty;
        var haystack = string.Join("\n", new[]
        {
            sheet.Title ?? string.Empty,
            sheet.Code ?? string.Empty,
            cachedCategory,
            ownerDisplay,
            role,
            sheet.Id ?? string.Empty,
        });

        return haystack.Contains(query, StringComparison.OrdinalIgnoreCase);
    }

    private void DrawSheetEditor()
    {
        this.ProcessPendingImportRequest();

        if (this.currentSheet is null || this.homeViewActive)
        {
            ImGui.PushStyleColor(ImGuiCol.Text, GrayDisabled);
            ImGui.TextWrapped("Open or create a sheet from the left panel. EZSheets keeps a local cache after you open a cloud sheet, so your latest opened sheets are still saved in your plugin config.");
            ImGui.PopStyleColor();
            ImGui.Spacing();
            this.DrawFakeBoldText("# Why do I have to Login?", AccentOrange);
            ImGui.Spacing();
            ImGui.TextWrapped("To be able to save your spread sheet without needing to actualy save it on your PC, we have our own database where everything is stored. By doing this way, you can also use the Share Code from each sheet to share with anyone you want and they will be able to see your sheet from the last time you saved it.");
            ImGui.Spacing();
            this.DrawFakeBoldText("# Why use Syncsheet?", AccentOrange);
            ImGui.Spacing();
            ImGui.TextWrapped(@"Creating/editing actual Excel Sheets can make your job on venues pretty slow. We all know how unnecessary complicated this can turn out if you are, for example, a Bartender with many customers or a manager with many staff and so on.

Syncsheet eliminates the slowness, the need for constant Alt+Tab and also completely eliminates the need of having a actual sheet in your PC and keep sending/tracking it to others. You can simply send your Share code and done, everything is shown in-game through a clean UI. Everything can be saved ingame without needing to actualy send the sheet back to anyone.

Sure, you could use a online shared google sheet but that still slows your job, you still need to be constantly alt+tabing and if you work/manage multiple venues, it can start to be a nightmare.

Syncsheet is here to help. The only Data we collect and save is the actual spread sheets so you are able to share them and save online. We do not collect any other type of data at all.
The discord Login method is the easyest to do this and the safest since it uses the oficial Discord login page.");
            return;
        }

        var presentationMode = this.currentSheet.Data.Settings.ViewMode == SheetViewMode.Presentation;
        if (!presentationMode)
        {
            this.DrawCurrentSheetToolbar();
            this.DrawSectionDivider();
            this.DrawSelectedCellValueRow();
            this.DrawSectionSpacer(1f, 0f);
            this.DrawSelectionFormatRow();
            this.DrawSectionSpacer(0f, 0f);
        }
        this.DrawSpreadsheetGrid();
        ImGui.SetCursorPosY(MathF.Max(0f, ImGui.GetCursorPosY() - 10f));
        this.DrawTabStrip();
    }

    private void DrawCurrentSheetToolbar()
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var toolbarButtonHeight = 22f;
        var displayedRoleLabel = this.GetDisplayedRoleLabel();
        var displayedRoleColor = this.GetDisplayedRoleColor();
        var displayedRoleTooltip = this.GetDisplayedRoleTooltip();

        ImGui.AlignTextToFramePadding();
        ImGui.TextUnformatted("Title");
        ImGui.SameLine(0, 6f);

        var title = StripAutoVersionSuffix(this.currentSheet.Title);
        var titleWidth = 240f;
        var canEditTitle = this.currentRole == SheetAccessRole.Owner;
        if (canEditTitle)
        {
            ImGui.SetNextItemWidth(titleWidth);
            if (ImGui.InputText("##SheetTitle", ref title, 128))
            {
                this.currentSheet.Title = StripAutoVersionSuffix(title);
                this.MarkPersistableSheetDirty();
            }
        }
        else
        {
            ImGui.BeginDisabled();
            ImGui.SetNextItemWidth(titleWidth);
            ImGui.InputText("##SheetTitleReadonly", ref title, 128, ImGuiInputTextFlags.ReadOnly);
            ImGui.EndDisabled();
            if (ImGui.IsItemHovered())
            {
                ImGui.SetTooltip("Only the Owner can change sheet name");
            }
        }

        ImGui.SameLine(0, 8f);
        var canDownloadFromCloud = this.currentRole != SheetAccessRole.Viewer;
        ImGui.BeginDisabled(!canDownloadFromCloud);
        if (this.DrawStyledButton("download-cloud", "Import from Cloud", AccentDownload, WhiteText, new Vector2(142f, toolbarButtonHeight)) && canDownloadFromCloud)
        {
            _ = this.RunActionAsync(async () => await this.LoadSheetAsync(this.currentSheet.Id).ConfigureAwait(false), "Cloud download failed.");
        }
        ImGui.EndDisabled();

        ImGui.SameLine(0, 8f);
        ImGui.BeginDisabled(!this.CanEditCurrentSheet);
        var saveCloudPressed = this.DrawStyledButton("save-cloud", "Save to Cloud", AccentSaveCloud, WhiteText, new Vector2(110f, toolbarButtonHeight));
        ImGui.EndDisabled();
        if (saveCloudPressed && this.CanEditCurrentSheet)
        {
            _ = this.RunActionAsync(async () => await this.SaveCurrentSheetToCloudAsync().ConfigureAwait(false), "Cloud save failed.");
        }

        if (this.currentRole == SheetAccessRole.Owner)
        {
            ImGui.SameLine(0, 8f);
            var snapshotLabel = DateTimeOffset.UtcNow < this.snapshotFeedbackUntilUtc ? "See History!" : "Snapshot";
            var snapshotPressed = this.DrawStyledButton("sheet-snapshot", snapshotLabel, HexToVector4("#6a3d8c"), WhiteText, new Vector2(88f, toolbarButtonHeight));
            if (snapshotPressed)
            {
                this.CreateSnapshot();
            }
            if (ImGui.IsItemHovered())
            {
                ImGui.SetTooltip("Create a restore snapshot for this sheet");
            }
        }

        var canDeleteSheet = this.CanDeleteCurrentSheet();
        var deletePressed = false;
        if (canDeleteSheet)
        {
            ImGui.SameLine(0, 8f);
            deletePressed = this.DrawStyledButton("delete-sheet", "Delete Sheet", AccentDelete, WhiteText, new Vector2(96f, toolbarButtonHeight));
            if (deletePressed)
            {
                ImGui.OpenPopup("Delete Sheet Confirmation Inline");
            }
        }

        if (ImGui.BeginPopup("Delete Sheet Confirmation Inline"))
        {
            ImGui.TextUnformatted("Delete current sheet?");
            if (this.DrawStyledButton("delete-inline-yes", "Yes", AccentDelete, WhiteText, new Vector2(58f, 22f)))
            {
                ImGui.CloseCurrentPopup();
                _ = this.RunActionAsync(async () =>
                {
                    if (this.currentSheet is null)
                    {
                        return;
                    }

                    var sheetId = this.currentSheet.Id;
                    await this.supabase.DeleteSheetAsync(sheetId).ConfigureAwait(false);
                    this.configuration.LocalSheets.Remove(sheetId);
                    this.configuration.SheetListColors.Remove(sheetId);
                    this.configuration.LastOpenedSheetId = string.Empty;
                    this.configuration.Save();
                    this.pendingClearedSheetId = sheetId;
                    this.pendingClearCurrentSheet = true;
                    this.statusMessage = "Sheet deleted.";
                }, "Could not delete sheet.");
            }

            ImGui.SameLine(0, 6f);
            if (this.DrawStyledButton("delete-inline-no", "No", HexToVector4("#666666"), WhiteText, new Vector2(52f, 22f)))
            {
                ImGui.CloseCurrentPopup();
            }

            ImGui.EndPopup();
        }

        var settings = this.GetSheetSettings();
        var favorite = this.IsSheetFavorited(this.currentSheet?.Id);
        settings.Favorite = favorite;
        var favoriteSize = new Vector2(28f, toolbarButtonHeight);
        var helpSize = new Vector2(70f, toolbarButtonHeight);
        var rolePrefix = "Role: ";
        var roleWidth = ImGui.CalcTextSize(rolePrefix).X + ImGui.CalcTextSize(displayedRoleLabel).X + ImGui.CalcTextSize($" | Size: {this.currentSheet.RowsCount} x {this.currentSheet.ColsCount}").X;
        var rightGroupWidth = roleWidth + 10f + favoriteSize.X + 8f + helpSize.X;
        var rightX = ImGui.GetWindowWidth() - rightGroupWidth - 22f;
        if (rightX > ImGui.GetCursorPosX())
        {
            ImGui.SameLine(rightX);
        }
        else
        {
            ImGui.SameLine();
        }

        ImGui.TextDisabled(rolePrefix);
        ImGui.SameLine(0, 0f);
        ImGui.TextColored(displayedRoleColor, displayedRoleLabel);
        if (ImGui.IsItemHovered() && !string.IsNullOrWhiteSpace(displayedRoleTooltip))
        {
            ImGui.SetTooltip(displayedRoleTooltip);
        }
        ImGui.SameLine(0, 0f);
        ImGui.TextDisabled($" | Size: {this.currentSheet.RowsCount} x {this.currentSheet.ColsCount}");
        ImGui.SameLine(0, 10f);
        if (this.DrawStyledButton("sheet-favorite", favorite ? "★" : "☆", favorite ? AccentOrange : HexToVector4("#5a5a5a"), WhiteText, favoriteSize))
        {
            favorite = !favorite;
            settings.Favorite = favorite;
            this.SetSheetFavoriteState(this.currentSheet?.Id, favorite);
            this.SaveCurrentSheetToLocalCache();
            this.configuration.Save();
        }
        if (ImGui.IsItemHovered())
        {
            ImGui.SetTooltip(settings.Favorite ? "Remove from favorites" : "Add to favorites");
        }
        ImGui.SameLine(0, 8f);
        if (this.DrawStyledButton("sheet-help", "Help", HexToVector4("#b218c9"), WhiteText, helpSize))
        {
            this.helpWindowOpen = true;
        }

        ImGui.NewLine();
        ImGui.SetCursorPosY(MathF.Max(0f, ImGui.GetCursorPosY() - 8f));
        if (this.currentRole == SheetAccessRole.Owner && this.currentSheet is not null)
        {
            var ownerSettings = this.GetSheetSettings();
            this.EnsurePermissionDefaults(ownerSettings);
            ownerSettings.TabPresets ??= new List<SheetTabPreset>();

            var orderedPresets = ownerSettings.TabPresets
                .OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (this.selectedSheetPresetIndex >= orderedPresets.Count)
            {
                this.selectedSheetPresetIndex = 0;
            }

            const float presetComboWidth = 170f;
            const float presetApplySpacing = 6f;
            const float presetApplyWidth = 24f;
            ImGui.AlignTextToFramePadding();
            ImGui.TextUnformatted("Sheet Preset");
            ImGui.SameLine(0, 6f);
            var presetComboStartX = ImGui.GetCursorPosX();
            var presetComboRightEdge = presetComboStartX + presetComboWidth + presetApplySpacing + presetApplyWidth;
            ImGui.BeginDisabled(orderedPresets.Count == 0);
            var presetNames = orderedPresets.Count == 0 ? new[] { "No presets saved" } : orderedPresets.Select(x => x.Name).ToArray();
            ImGui.SetNextItemWidth(presetComboWidth);
            ImGui.Combo("##SheetPresetCombo", ref this.selectedSheetPresetIndex, presetNames, presetNames.Length);
            ImGui.SameLine(0, presetApplySpacing);
            var applyPresetPressed = this.DrawStyledButton("apply-sheet-preset", "✓", HexToVector4("#2a7804"), WhiteText, new Vector2(presetApplyWidth, 22f));
            ImGui.EndDisabled();
            if (applyPresetPressed && orderedPresets.Count > 0)
            {
                this.CreateTabFromSavedPreset(orderedPresets[Math.Clamp(this.selectedSheetPresetIndex, 0, orderedPresets.Count - 1)]);
            }
            if (ImGui.IsItemHovered())
            {
                ImGui.SetTooltip("Create a new tab from the selected preset");
            }

            var defaultRoleItems = ownerSettings.PermissionPresets
                .OrderBy(preset => string.Equals(preset.Name, "Viewer", StringComparison.OrdinalIgnoreCase) ? 0 : string.Equals(preset.Name, "Editor", StringComparison.OrdinalIgnoreCase) ? 1 : 2)
                .ThenBy(preset => preset.Name, StringComparer.OrdinalIgnoreCase)
                .Select(preset => preset.Name)
                .ToList();
            if (defaultRoleItems.Count == 0)
            {
                defaultRoleItems.AddRange(new[] { "Viewer", "Editor" });
            }

            var currentDefaultRole = this.currentSheet.DefaultRole;
            var defaultRoleIndex = defaultRoleItems.FindIndex(item => string.Equals(item, currentDefaultRole, StringComparison.OrdinalIgnoreCase));
            if (defaultRoleIndex < 0 || defaultRoleIndex >= defaultRoleItems.Count)
            {
                defaultRoleItems.Insert(0, string.IsNullOrWhiteSpace(currentDefaultRole) ? "Viewer" : currentDefaultRole);
                defaultRoleIndex = 0;
            }

            ImGui.AlignTextToFramePadding();
            ImGui.TextUnformatted("Default Join Role");
            ImGui.SameLine(0, 6f);
            var defaultRoleComboStartX = ImGui.GetCursorPosX();
            var defaultRoleComboWidth = MathF.Max(90f, presetComboRightEdge - defaultRoleComboStartX);
            ImGui.SetNextItemWidth(defaultRoleComboWidth);
            if (ImGui.Combo("##DefaultJoinRole", ref defaultRoleIndex, defaultRoleItems.ToArray(), defaultRoleItems.Count))
            {
                this.currentSheet.DefaultRole = defaultRoleItems[defaultRoleIndex];
            }
        }

        var canUseInviteCodes = this.CanUseInviteCodes();
        if (this.currentRole == SheetAccessRole.Owner || canUseInviteCodes)
        {
            this.EnsureCurrentUniqueCodeAvailable();

            var now = DateTimeOffset.UtcNow;
            const float copyButtonWidth = 104f;
            const float smallButtonWidth = 24f;
            const float hourButtonWidth = 46f;
            const float minuteButtonWidth = 46f;
            const float rowHeight = 22f;
            const float rowSpacing = 4f;
            const float itemGap = 4f;
            var shareColor = HexToVector4("#de7300");
            var uniqueColor = HexToVector4("#20a83c");
            var temporaryColor = HexToVector4("#6bc7ff");
            var shareCode = string.IsNullOrWhiteSpace(this.currentSheet?.Code) ? "--------" : this.currentSheet!.Code;
            var displayedShareCode = this.hideSharedCode ? "**********" : shareCode;
            var uniqueCode = this.currentOwnerUniqueCode?.Code ?? "--------";
            var temporaryCode = this.GetLatestActiveTemporaryInviteCode()?.Code ?? this.cachedLatestTemporaryInviteCode ?? "--------";
            var temporaryDisplay = now < this.temporaryCodeCopiedFeedbackUntilUtc ? "Copied!..." : temporaryCode;
            var codeBoxWidth = new[] { displayedShareCode, uniqueCode, temporaryDisplay, "**********", "Copied!..." }
                .Select(value => ImGui.CalcTextSize(value).X)
                .DefaultIfEmpty(96f)
                .Max() + 16f;
            codeBoxWidth = MathF.Max(110f, codeBoxWidth);

            var startX = ImGui.GetCursorPosX();
            var startY = ImGui.GetCursorPosY() + 6f;
            var drawList = ImGui.GetWindowDrawList();

            void DrawCodeBox(string id, string value, Vector4 textColor, float x, float y, string? tooltip = null)
            {
                var min = ImGui.GetWindowPos() + new Vector2(x, y);
                var max = min + new Vector2(codeBoxWidth, rowHeight);
                drawList.AddRectFilled(min, max, ImGui.ColorConvertFloat4ToU32(new Vector4(0.18f, 0.18f, 0.18f, 0.45f)), 3f);
                drawList.AddText(min + new Vector2(8f, 3f), ImGui.ColorConvertFloat4ToU32(textColor), value);
                ImGui.SetCursorPos(new Vector2(x, y));
                ImGui.InvisibleButton(id, new Vector2(codeBoxWidth, rowHeight));
                if (!string.IsNullOrWhiteSpace(tooltip) && ImGui.IsItemHovered())
                {
                    ImGui.SetTooltip(tooltip);
                }
            }

            var rowIndex = 0;
            float RowY(int index) => startY + (index * (rowHeight + rowSpacing));

            if (this.currentRole == SheetAccessRole.Owner)
            {
                var rowY = RowY(rowIndex);
                var copySharedLabel = now < this.copyCodeFeedbackUntilUtc ? "Copied!" : "Copy Shared";
                ImGui.SetCursorPos(new Vector2(startX, rowY));
                if (this.DrawStyledButton("copy-code", copySharedLabel, shareColor, WhiteText, new Vector2(copyButtonWidth, rowHeight)))
                {
                    ImGui.SetClipboardText(shareCode);
                    this.copyCodeFeedbackUntilUtc = DateTimeOffset.UtcNow.AddSeconds(3);
                    this.statusMessage = "Share code copied to clipboard.";
                }

                var visibilityX = startX + copyButtonWidth + itemGap;
                ImGui.SetCursorPos(new Vector2(visibilityX, rowY));
                var visibilityLabel = this.hideSharedCode ? "●" : "○";
                if (this.DrawStyledButton("toggle-share-code-visibility", visibilityLabel, shareColor, WhiteText, new Vector2(smallButtonWidth, rowHeight)))
                {
                    this.hideSharedCode = !this.hideSharedCode;
                }
                if (ImGui.IsItemHovered())
                {
                    ImGui.SetTooltip(this.hideSharedCode ? "Show shared code" : "Hide shared code");
                }

                var codeX = visibilityX + smallButtonWidth + itemGap;
                DrawCodeBox("shared-code-box", displayedShareCode, shareColor, codeX, rowY, "This code never expires, share with caution");
                rowIndex++;
            }

            if (canUseInviteCodes || this.currentRole == SheetAccessRole.Owner)
            {
                var rowY = RowY(rowIndex);
                var copyUniqueLabel = now < this.copyUniqueFeedbackUntilUtc ? "Copied!" : "Copy Unique";
                ImGui.SetCursorPos(new Vector2(startX, rowY));
                if (this.DrawStyledButton("copy-unique-code", copyUniqueLabel, uniqueColor, WhiteText, new Vector2(copyButtonWidth, rowHeight)))
                {
                    var copiedUniqueCode = this.currentOwnerUniqueCode?.Code ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(copiedUniqueCode))
                    {
                        ImGui.SetClipboardText(copiedUniqueCode);
                        this.copyUniqueFeedbackUntilUtc = DateTimeOffset.UtcNow.AddSeconds(3);
                        this.statusMessage = "Unique code copied to clipboard.";
                        _ = Task.Run(async () =>
                        {
                            try
                            {
                                if (this.currentSheet is null)
                                {
                                    return;
                                }

                                var nextCode = await this.supabase.CreateUniqueCodeAsync(this.currentSheet.Id, invalidateCurrent: false).ConfigureAwait(false);
                                if (this.currentSheet is not null)
                                {
                                    this.currentOwnerUniqueCode = nextCode;
                                    this.uniqueCodeEnsureSheetId = this.currentSheet.Id;
                                    this.AppendLocalInviteAuditEntry(nextCode.Code, "unique", 0, null, true);
                                }
                            }
                            catch
                            {
                            }
                        });
                    }
                    else
                    {
                        this.statusMessage = "Unique code is still loading.";
                    }
                }

                var refreshX = startX + copyButtonWidth + itemGap;
                ImGui.SetCursorPos(new Vector2(refreshX, rowY));
                if (this.DrawStyledButton("refresh-unique-code", "◎", uniqueColor, WhiteText, new Vector2(smallButtonWidth, rowHeight)))
                {
                    _ = this.RunActionAsync(async () =>
                    {
                        if (this.currentSheet is null)
                        {
                            return;
                        }

                        this.currentOwnerUniqueCode = await this.supabase.CreateUniqueCodeAsync(this.currentSheet.Id, invalidateCurrent: true).ConfigureAwait(false);
                        this.uniqueCodeEnsureSheetId = this.currentSheet.Id;
                        this.AppendLocalInviteAuditEntry(this.currentOwnerUniqueCode.Code, "unique", 0, null, true);
                        this.statusMessage = "A new unique code is now active.";
                    }, "Could not refresh the unique code.");
                }
                if (ImGui.IsItemHovered())
                {
                    ImGui.SetTooltip("Generate new unique code");
                }

                var codeX = refreshX + smallButtonWidth + itemGap;
                DrawCodeBox("unique-code-box", uniqueCode, uniqueColor, codeX, rowY, "This code becomes invalid after 1 use");
                rowIndex++;
            }

            if (this.currentRole == SheetAccessRole.Owner || canUseInviteCodes)
            {
                var rowY = RowY(rowIndex);
                var hourLabel = this.temporaryInviteHoursSelection > 0 ? this.temporaryInviteHoursSelection.ToString("00", CultureInfo.InvariantCulture) : string.Empty;
                var hourButtonText = string.IsNullOrWhiteSpace(hourLabel) ? "Hour" : hourLabel;
                var minuteLabel = this.temporaryInviteMinutesSelection > 0 ? this.temporaryInviteMinutesSelection.ToString("00", CultureInfo.InvariantCulture) : string.Empty;
                var minuteButtonText = string.IsNullOrWhiteSpace(minuteLabel) ? "Mins" : minuteLabel;
                var colonWidth = ImGui.CalcTextSize(":").X;
                var generateX = startX + copyButtonWidth + itemGap;
                var minuteX = generateX - itemGap - minuteButtonWidth;
                var colonX = minuteX - itemGap - colonWidth;
                var hourX = colonX - itemGap - hourButtonWidth;

                ImGui.SetCursorPos(new Vector2(hourX, rowY));
                if (this.DrawStyledButton("temporary-hours-button", hourButtonText, temporaryColor, string.IsNullOrWhiteSpace(hourLabel) ? GrayDisabled : WhiteText, new Vector2(hourButtonWidth, rowHeight)))
                {
                    ImGui.OpenPopup("temporary-hours-popup");
                }
                if (ImGui.IsItemHovered())
                {
                    ImGui.SetTooltip("Hours");
                }
                if (ImGui.BeginPopup("temporary-hours-popup"))
                {
                    for (var hour = 1; hour <= 12; hour++)
                    {
                        var option = hour.ToString("00", CultureInfo.InvariantCulture);
                        if (ImGui.Selectable(option, hour == this.temporaryInviteHoursSelection))
                        {
                            this.temporaryInviteHoursSelection = hour;
                        }
                    }
                    ImGui.EndPopup();
                }

                ImGui.SetCursorPos(new Vector2(colonX, rowY + 3.5f));
                ImGui.AlignTextToFramePadding();
                ImGui.TextColored(temporaryColor, ":");

                ImGui.SetCursorPos(new Vector2(minuteX, rowY));
                if (this.DrawStyledButton("temporary-minutes-button", minuteButtonText, temporaryColor, string.IsNullOrWhiteSpace(minuteLabel) ? GrayDisabled : WhiteText, new Vector2(minuteButtonWidth, rowHeight)))
                {
                    ImGui.OpenPopup("temporary-minutes-popup");
                }
                if (ImGui.IsItemHovered())
                {
                    ImGui.SetTooltip("Minutes");
                }
                if (ImGui.BeginPopup("temporary-minutes-popup"))
                {
                    if (ImGui.BeginChild("##temporary-minutes-scroll", new Vector2(72f, 13f * 22f), false, ImGuiWindowFlags.AlwaysVerticalScrollbar))
                    {
                        for (var minute = 0; minute <= 59; minute++)
                        {
                            var option = minute.ToString("00", CultureInfo.InvariantCulture);
                            if (ImGui.Selectable(option, minute == this.temporaryInviteMinutesSelection))
                            {
                                this.temporaryInviteMinutesSelection = minute;
                            }
                        }
                    }
                    ImGui.EndChild();
                    ImGui.EndPopup();
                }

                ImGui.SetCursorPos(new Vector2(generateX, rowY));
                if (this.DrawStyledButton("generate-temporary-code", "✓", temporaryColor, WhiteText, new Vector2(smallButtonWidth, rowHeight)))
                {
                    _ = this.RunActionAsync(this.GenerateTemporaryInviteCodeAsync, "Could not generate the temporary access code.");
                }
                if (ImGui.IsItemHovered())
                {
                    ImGui.SetTooltip("Generate and copy a one-time temporary access code");
                }

                var codeX = generateX + smallButtonWidth + itemGap;
                DrawCodeBox("temporary-code-box", temporaryDisplay, temporaryColor, codeX, rowY, "One-time access code with a custom duration");
                rowIndex++;
            }

            ImGui.SetCursorPosY(startY + (rowIndex * (rowHeight + rowSpacing)) + 2f);
        }
    }

    private TemporaryInviteCode? GetLatestActiveTemporaryInviteCode()
    {
        if (this.currentSheet is null)
        {
            return null;
        }

        var settings = this.GetSheetSettings();
        settings.TemporaryInviteCodes ??= new List<TemporaryInviteCode>();
        return settings.TemporaryInviteCodes
            .Where(code => !code.Invalidated && !code.UsedAtUtc.HasValue)
            .OrderByDescending(code => code.CreatedAtUtc)
            .FirstOrDefault();
    }

    private async Task GenerateTemporaryInviteCodeAsync()
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var settings = this.GetSheetSettings();
        settings.TemporaryInviteCodes ??= new List<TemporaryInviteCode>();
        settings.InviteAuditLog ??= new List<InviteAuditEntry>();

        var durationMinutes = (Math.Clamp(this.temporaryInviteHoursSelection, 1, 12) * 60) + Math.Clamp(this.temporaryInviteMinutesSelection, 0, 59);
        durationMinutes = Math.Max(1, durationMinutes);

        var createdByName = this.GetSafeCharacterFullName();
        var code = $"T{SupabaseRestClient.GenerateShareCode(9)}";
        while (settings.TemporaryInviteCodes.Any(existing => string.Equals(existing.Code, code, StringComparison.OrdinalIgnoreCase)))
        {
            code = $"T{SupabaseRestClient.GenerateShareCode(9)}";
        }

        this.cachedLatestTemporaryInviteCode = code;
        var entry = new TemporaryInviteCode
        {
            Id = Guid.NewGuid().ToString("N"),
            Code = code,
            CreatedByUserId = this.configuration.UserId,
            CreatedByName = createdByName,
            CreatedAtUtc = DateTimeOffset.UtcNow,
            DurationMinutes = durationMinutes,
        };
        settings.TemporaryInviteCodes.Add(entry);
        while (settings.TemporaryInviteCodes.Count > 100)
        {
            settings.TemporaryInviteCodes.RemoveAt(0);
        }

        settings.InviteAuditLog.Add(new InviteAuditEntry
        {
            Id = Guid.NewGuid().ToString("N"),
            Code = code,
            CodeType = "temporary",
            CreatedByUserId = this.configuration.UserId,
            CreatedByName = createdByName,
            CreatedAtUtc = entry.CreatedAtUtc,
            DurationMinutes = durationMinutes,
            WasUsed = false,
        });
        while (settings.InviteAuditLog.Count > 300)
        {
            settings.InviteAuditLog.RemoveAt(0);
        }

        this.RecordActivity("Generated temporary access code", null, null, code);

        try
        {
            ImGui.SetClipboardText(code);
            this.temporaryCodeCopiedFeedbackUntilUtc = DateTimeOffset.UtcNow.AddSeconds(3);
        }
        catch
        {
        }

        await this.SaveCurrentSheetToCloudAsync().ConfigureAwait(false);
        this.statusMessage = $"Temporary access code generated for {this.FormatDurationMinutes(durationMinutes)} and copied to clipboard.";
    }

    private void AppendLocalInviteAuditEntry(string code, string codeType, int durationMinutes, string? createdByName = null, bool saveImmediately = false)
    {
        if (this.currentSheet is null || string.IsNullOrWhiteSpace(code))
        {
            return;
        }

        var settings = this.GetSheetSettings();
        settings.InviteAuditLog ??= new List<InviteAuditEntry>();
        var normalizedCode = code.Trim().ToUpperInvariant();
        var normalizedType = codeType.Trim().ToLowerInvariant();
        if (settings.InviteAuditLog.Any(entry => string.Equals(entry.Code?.Trim(), normalizedCode, StringComparison.OrdinalIgnoreCase)
            && string.Equals(entry.CodeType?.Trim(), normalizedType, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        settings.InviteAuditLog.Add(new InviteAuditEntry
        {
            Id = Guid.NewGuid().ToString("N"),
            Code = normalizedCode,
            CodeType = normalizedType,
            CreatedByUserId = this.configuration.UserId,
            CreatedByName = string.IsNullOrWhiteSpace(createdByName) ? this.GetSafeCharacterFullName() : createdByName!,
            CreatedAtUtc = DateTimeOffset.UtcNow,
            DurationMinutes = Math.Max(0, durationMinutes),
            WasUsed = false,
        });

        while (settings.InviteAuditLog.Count > 300)
        {
            settings.InviteAuditLog.RemoveAt(0);
        }

        if (saveImmediately)
        {
            this.SaveCurrentSheetToLocalCache();
        }
    }

    private void EnsureCurrentUniqueCodeAvailable()
    {
        if (this.currentSheet is null || !(this.currentRole == SheetAccessRole.Owner || this.CanUseInviteCodes()))
        {
            this.uniqueCodeEnsureSheetId = null;
            this.uniqueCodeEnsureInFlight = false;
            return;
        }

        if (this.currentOwnerUniqueCode is not null && string.Equals(this.uniqueCodeEnsureSheetId, this.currentSheet.Id, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        if (this.uniqueCodeEnsureInFlight)
        {
            return;
        }

        this.uniqueCodeEnsureInFlight = true;
        var sheetId = this.currentSheet.Id;
        _ = Task.Run(async () =>
        {
            try
            {
                var code = await this.supabase.CreateUniqueCodeAsync(sheetId, invalidateCurrent: false).ConfigureAwait(false);
                if (this.currentSheet is not null && string.Equals(this.currentSheet.Id, sheetId, StringComparison.OrdinalIgnoreCase))
                {
                    this.currentOwnerUniqueCode = code;
                    this.uniqueCodeEnsureSheetId = sheetId;
                }
            }
            catch
            {
                // Keep the UI responsive even if the code cannot be prefetched yet.
            }
            finally
            {
                this.uniqueCodeEnsureInFlight = false;
            }
        });
    }

    private void DrawTabStrip()
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var document = this.currentSheet.Data;
        document.Normalize(this.currentSheet.RowsCount, this.currentSheet.ColsCount);

        ImGui.AlignTextToFramePadding();
        var tabsLineStartX = ImGui.GetCursorPosX();
        ImGui.TextUnformatted("Tabs:");
        ImGui.SameLine(tabsLineStartX + 44f, 6f);

        for (var index = 0; index < document.Tabs.Count; index++)
        {
            if (index > 0)
            {
                ImGui.SameLine(0, 4f);
            }

            var isActive = document.ActiveTabIndex == index;
            var tab = document.Tabs[index];
            var canDeleteTabs = this.CanCreateTabs() && document.Tabs.Count > 1;
            var closeWidth = canDeleteTabs ? 20f : 0f;
            var width = MathF.Max(72f, ImGui.CalcTextSize(tab.Name).X + 24f);
            var tabBaseColor = tab.TabColor != 0 ? ImGui.ColorConvertU32ToFloat4(tab.TabColor) : (isActive ? AccentOrange : HexToVector4("#2b2b2b"));
            var color = isActive ? tabBaseColor : AdjustBrightness(tabBaseColor, 0.68f);
            var textColor = isActive ? WhiteText : HexToVector4("#d0d0d0");
            ImGui.BeginGroup();
            if (this.DrawStyledButton($"tab-{index}", tab.Name, color, textColor, new Vector2(width, 22)))
            {
                if (isActive && this.CanCreateTabs())
                {
                    ImGui.OpenPopup($"rename-tab-{index}");
                }
                else
                {
                    document.ActiveTabIndex = index;
                    this.EnsureSelectionInBounds();
                }
            }
            if (canDeleteTabs)
            {
                ImGui.SameLine(0f, 0f);
                if (this.DrawStyledButton($"tab-delete-{index}", "X", color, textColor, new Vector2(closeWidth, 22f)))
                {
                    this.pendingTabDeletePopupPositions[$"delete-tab-popup-{index}"] = new Vector2(ImGui.GetItemRectMin().X - 62f, ImGui.GetItemRectMin().Y - 58f);
                    ImGui.OpenPopup($"delete-tab-popup-{index}");
                }
                if (ImGui.IsItemHovered())
                {
                    ImGui.SetTooltip($"Delete {tab.Name}");
                }

                if (this.pendingTabDeletePopupPositions.TryGetValue($"delete-tab-popup-{index}", out var popupPosition))
                {
                    ImGui.SetNextWindowPos(popupPosition, ImGuiCond.Appearing);
                }

                if (ImGui.BeginPopup($"delete-tab-popup-{index}"))
                {
                    ImGui.TextUnformatted($"Delete {tab.Name}?");
                    if (this.DrawStyledButton($"tab-delete-confirm-{index}", "Yes", AccentDelete, WhiteText, new Vector2(52f, 22f)))
                    {
                        this.PushUndoSnapshot();
                        var deletedTabName = tab.Name;
                        var activeTabName = this.GetActiveTab()?.Name;
                        document.Tabs.RemoveAt(index);
                        if (document.ActiveTabIndex >= document.Tabs.Count)
                        {
                            document.ActiveTabIndex = Math.Max(0, document.Tabs.Count - 1);
                        }
                        else if (document.ActiveTabIndex > index)
                        {
                            document.ActiveTabIndex--;
                        }

                        if (!string.IsNullOrWhiteSpace(activeTabName) && string.Equals(activeTabName, deletedTabName, StringComparison.OrdinalIgnoreCase))
                        {
                            this.ReleaseCellLockIfNeeded(this.editingCellKey);
                            this.editingCellKey = null;
                            this.editingCellBuffer = string.Empty;
                            this.editingFocusRequested = false;
                        }

                        this.pendingTabDeletePopupPositions.Remove($"delete-tab-popup-{index}");
                        this.ClearSelection();
                        this.SelectSingleCell("R1C1");
                        ImGui.CloseCurrentPopup();
                        if (this.CanPersistCurrentSheetChanges())
                        {
                            this.RequestBackgroundCloudSave("Could not sync tab deletion.");
                        }
                    }

                    ImGui.SameLine(0f, 6f);
                    if (this.DrawStyledButton($"tab-delete-cancel-{index}", "No", HexToVector4("#666666"), WhiteText, new Vector2(48f, 22f)))
                    {
                        this.pendingTabDeletePopupPositions.Remove($"delete-tab-popup-{index}");
                        ImGui.CloseCurrentPopup();
                    }

                    ImGui.EndPopup();
                }
            }
            ImGui.EndGroup();

            if (ImGui.BeginPopup($"rename-tab-{index}"))
            {
                var rename = tab.Name;
                ImGui.SetNextItemWidth(180f);
                if (ImGui.InputText($"##rename-tab-input-{index}", ref rename, 128, ImGuiInputTextFlags.EnterReturnsTrue))
                {
                    this.PushUndoSnapshot();
                    tab.Name = string.IsNullOrWhiteSpace(rename) ? $"Sheet {index + 1}" : rename.Trim();
                    ImGui.CloseCurrentPopup();
                }
                var tabColor = tab.TabColor == 0 ? color : ImGui.ColorConvertU32ToFloat4(tab.TabColor);
                if (ImGui.ColorEdit4($"##tab-color-{index}", ref tabColor, ImGuiColorEditFlags.NoInputs | ImGuiColorEditFlags.NoLabel | ImGuiColorEditFlags.AlphaBar))
                {
                    tab.TabColor = ImGui.ColorConvertFloat4ToU32(tabColor);
                }
                if (ImGui.IsItemHovered())
                {
                    ImGui.SetTooltip("Tab color");
                }
                ImGui.SameLine(0f, 6f);
                if (ImGui.SmallButton($"Reset##tab-color-reset-{index}"))
                {
                    tab.TabColor = 0;
                }
                ImGui.TextDisabled("Press Enter to save the new tab name.");
                ImGui.EndPopup();
            }
        }

        ImGui.SameLine(0, 4f);
        ImGui.BeginDisabled(!this.CanCreateTabs());
        var addTabPressed = this.DrawStyledButton("tab-add", "+", AccentTabAdd, WhiteText, new Vector2(22, 22));
        ImGui.EndDisabled();
        if (addTabPressed && this.CanCreateTabs())
        {
            this.PushUndoSnapshot();
            document.AddNewTab();
            this.pendingTabDeletePopupPositions.Clear();
            this.ClearSelection();
            this.SelectSingleCell("R1C1");
        }

        ImGui.SameLine(0, 4f);
        ImGui.BeginDisabled(!this.CanCreateTabs() || document.Tabs.Count == 0);
        var duplicatePressed = this.DrawStyledButton("tab-duplicate", "D", HexToVector4("#496a9d"), WhiteText, new Vector2(22, 22));
        ImGui.EndDisabled();
        if (duplicatePressed && this.CanCreateTabs() && document.Tabs.Count > 0)
        {
            this.DuplicateActiveTab();
        }
        if (ImGui.IsItemHovered())
        {
            ImGui.SetTooltip("Duplicate active tab");
        }
    }

    private void DrawSelectedCellValueRow()
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var sheet = this.currentSheet;
        var activeTab = this.GetActiveTab();
        if (sheet is null || activeTab is null)
        {
            return;
        }

        var primaryKey = this.GetPrimarySelectedCellKey();
        var styleSourceCell = activeTab.GetOrCreateCell(primaryKey);
        activeTab.TryGetFeature(primaryKey, out var selectedFeature);

        ImGui.TextUnformatted("Selected Cell:");
        ImGui.SameLine(0, 8f);
        ImGui.PushStyleColor(ImGuiCol.Text, GrayDisabled);
        ImGui.TextUnformatted("Double-click a cell to edit it directly. CTRL + click adds more cells to the selection");
        ImGui.PopStyleColor();

        var canOpenPermissions = this.CanOpenPermissionsPanel();
        var canOpenBlocklist = this.currentSheet is not null && this.currentRole == SheetAccessRole.Owner;
        if (canOpenPermissions || canOpenBlocklist)
        {
            var permissionsSize = new Vector2(96f, 22f);
            var blocklistSize = new Vector2(84f, 22f);
            var totalWidth = (canOpenBlocklist ? blocklistSize.X : 0f) + (canOpenPermissions ? permissionsSize.X : 0f) + ((canOpenBlocklist && canOpenPermissions) ? 6f : 0f);
            var buttonsX = ImGui.GetWindowContentRegionMax().X - totalWidth;
            if (buttonsX > ImGui.GetCursorPosX())
            {
                ImGui.SameLine(buttonsX);
            }
            else
            {
                ImGui.SameLine();
            }

            if (canOpenBlocklist)
            {
                if (this.DrawStyledButton("sheet-blocklist-open", "Blocklist", HexToVector4("#8a2043"), WhiteText, blocklistSize))
                {
                    this.OpenBlocklistWindow();
                }
                if (canOpenPermissions)
                {
                    ImGui.SameLine(0, 6f);
                }
            }

            if (canOpenPermissions && this.DrawStyledButton("sheet-permissions-open", "Permissions", HexToVector4("#d46728"), WhiteText, permissionsSize))
            {
                this.OpenPermissionsWindow();
            }
        }

        var avail = ImGui.GetContentRegionAvail().X;
        var leftWidth = MathF.Max(260f, avail);

        if (this.selectedCellKeys.Count == 1)
        {
            var value = styleSourceCell.Value;
            ImGui.SetNextItemWidth(leftWidth);
            if (selectedFeature?.Dropdown is not null && selectedFeature.Dropdown.Options.Count > 0)
            {
                var options = selectedFeature.Dropdown.Options;
                var selectedIndex = Math.Max(0, options.FindIndex(x => string.Equals(x, value, StringComparison.OrdinalIgnoreCase)));
                ImGui.SetNextItemWidth(leftWidth);
                if (ImGui.Combo("##CellDropdownValue", ref selectedIndex, options.ToArray(), options.Count))
                {
                    this.PushUndoSnapshot();
                    styleSourceCell.Value = options[selectedIndex];
                }
            }
            else if (this.CanEditCurrentSheet)
            {
                if (ImGui.InputText("##CellValue", ref value, 4096))
                {
                    this.PushUndoSnapshot();
                    styleSourceCell.Value = value;
                }
            }
            else
            {
                ImGui.InputText("##CellValueReadonly", ref value, 4096, ImGuiInputTextFlags.ReadOnly);
            }
        }
        else
        {
            ImGui.BeginDisabled();
            var summary = string.Join(", ", this.selectedCellKeys.Select(ToA1).OrderBy(s => s).Take(4)) + (this.selectedCellKeys.Count > 4 ? "..." : string.Empty);
            ImGui.SetNextItemWidth(leftWidth);
            ImGui.InputText("##CellValueMulti", ref summary, 4096, ImGuiInputTextFlags.ReadOnly);
            ImGui.EndDisabled();
        }
    }

    private void DrawSelectionFormatRow()
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var sheet = this.currentSheet;
        var activeTab = this.GetActiveTab();
        if (sheet is null || activeTab is null)
        {
            return;
        }

        const float saveLocalWidth = 90f;
        const float importWidth = 102f;
        const float rowButtonHeight = 22f;
        var selectionCount = this.selectedCellKeys.Count;
        var primaryKey = this.GetPrimarySelectedCellKey();
        var styleSourceCell = activeTab.GetOrCreateCell(primaryKey);
        var cellButtonLabel = selectionCount <= 1 ? ToA1(primaryKey) : string.Join(", ", this.selectedCellKeys.Select(ToA1).OrderBy(s => s).Take(3)) + (selectionCount > 3 ? "..." : string.Empty);

        if (this.DrawStyledButton("undo-change", "←", HexToVector4("#3b3b3b"), WhiteText, new Vector2(30, rowButtonHeight)))
        {
            this.UndoLastChange();
        }

        ImGui.SameLine(0, 4f);
        if (this.DrawStyledButton("redo-change-inline", "→", HexToVector4("#3b3b3b"), WhiteText, new Vector2(30, rowButtonHeight)))
        {
            this.RedoLastChange();
        }

        ImGui.SameLine(0, 4f);
        this.DrawStyledButton("selected-cell-indicator", cellButtonLabel, HexToVector4("#3b3b3b"), WhiteText, new Vector2(90, rowButtonHeight));

        ImGui.SameLine(0, 6f);
        if (this.CanEditCurrentSheet)
        {
            this.DrawCompactFormatToolbar(activeTab, primaryKey, styleSourceCell);
        }
        else
        {
            ImGui.PushStyleColor(ImGuiCol.Text, GrayDisabled);
            ImGui.TextWrapped("This sheet is read-only for your account.");
            ImGui.PopStyleColor();
        }

        const float findWidth = 64f;
        const float historyWidth = 78f;
        var rightBlockWidth = findWidth + 8f + historyWidth + 10f + saveLocalWidth + 8f + importWidth;
        var rightBlockStartX = ImGui.GetWindowContentRegionMax().X - rightBlockWidth;
        if (rightBlockStartX > ImGui.GetCursorPosX())
        {
            ImGui.SameLine(rightBlockStartX);
        }
        else
        {
            ImGui.SameLine();
        }

        if (this.DrawStyledButton("find-action-inline", "Find", HexToVector4("#355f8c"), WhiteText, new Vector2(findWidth, rowButtonHeight)))
        {
            this.searchWindowOpen = true;
            this.pendingSearchWindowPosition = new Vector2(ImGui.GetItemRectMin().X, ImGui.GetItemRectMax().Y + 6f);
            this.pendingSearchWindowReposition = true;
            this.BuildSearchResults();
        }

        ImGui.SameLine(0, 8f);
        var canSeeHistory = this.HasSheetPermission(SheetPermissionType.SeeHistory);
        ImGui.BeginDisabled(!canSeeHistory);
        if (this.DrawStyledButton("history-action-inline", "History", HexToVector4("#6a3d8c"), WhiteText, new Vector2(historyWidth, rowButtonHeight)) && canSeeHistory)
        {
            this.historyWindowOpen = true;
        }
        ImGui.EndDisabled();

        var canSaveLocal = this.HasSheetPermission(SheetPermissionType.SaveLocal);
        ImGui.SameLine(0, 10f);
        ImGui.BeginDisabled(!canSaveLocal);
        if (this.DrawStyledButton("save-local", "Save Local", AccentSaveLocal, WhiteText, new Vector2(saveLocalWidth, rowButtonHeight)) && canSaveLocal)
        {
            try
            {
                this.SaveCurrentSheetToLocalCache();
                this.lastSavedLocalPath = this.ExportCurrentSheetToLocalFile();
                this.localSavePopupOpen = true;
                this.localSavePopupRequested = true;
                this.statusMessage = "Saved the current sheet locally.";
            }
            catch (Exception ex)
            {
                this.statusMessage = ex.Message;
            }
        }
        ImGui.EndDisabled();

        var canImportSheet = this.HasSheetPermission(SheetPermissionType.ImportSheet);
        ImGui.SameLine(0, 8f);
        ImGui.BeginDisabled(!canImportSheet);
        if (this.DrawStyledButton("import-sheet", "Import Sheet", HexToVector4("#d4b013"), WhiteText, new Vector2(importWidth, rowButtonHeight)) && canImportSheet)
        {
            this.ImportSheetFromFile();
        }
        if (ImGui.IsItemHovered())
        {
            ImGui.BeginTooltip();
            this.DrawFakeBoldText("Experimental feature", HexToVector4("#ff9b9b"));
            ImGui.EndTooltip();
        }
        ImGui.EndDisabled();
    }

    private void DrawSpreadsheetGrid()
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var sheet = this.currentSheet;
        var activeTab = this.GetActiveTab();
        if (sheet is null || activeTab is null)
        {
            return;
        }

        const float rowHeaderWidth = 34f;
        const float defaultColumnWidth = 140f;
        const float defaultCellHeight = 24f;
        var flags = ImGuiTableFlags.ScrollX | ImGuiTableFlags.ScrollY | ImGuiTableFlags.SizingFixedFit | ImGuiTableFlags.NoSavedSettings;

        const float bottomTabReserve = 30f;
        var gridHostSize = ImGui.GetContentRegionAvail();
        gridHostSize.Y = MathF.Max(120f, gridHostSize.Y - bottomTabReserve);
        if (!ImGui.BeginChild("##SheetGridHost", gridHostSize, false, ImGuiWindowFlags.NoScrollbar | ImGuiWindowFlags.NoScrollWithMouse))
        {
            ImGui.EndChild();
            return;
        }

        this.gridClipMin = ImGui.GetWindowPos() + ImGui.GetWindowContentRegionMin();
        this.gridClipMax = ImGui.GetWindowPos() + ImGui.GetWindowContentRegionMax();
        this.pendingMergedTextOverlays.Clear();
        this.pendingMergedSelectionOverlays.Clear();
        this.pendingMergedEditOverlay = null;

        var tableSize = ImGui.GetContentRegionAvail();
        tableSize.Y = MathF.Max(120f, tableSize.Y);
        ImGui.PushStyleVar(ImGuiStyleVar.CellPadding, Vector2.Zero);
        if (!ImGui.BeginTable("##SheetGrid", sheet.ColsCount + 1, flags, tableSize))
        {
            ImGui.PopStyleVar();
            ImGui.EndChild();
            return;
        }

        var freezeRows = 1 + Math.Max(0, this.GetSheetSettings().FrozenRows);
        var freezeColumns = 1 + Math.Max(0, this.GetSheetSettings().FrozenColumns);
        ImGui.TableSetupScrollFreeze(freezeColumns, freezeRows);
        ImGui.TableSetupColumn("##corner", ImGuiTableColumnFlags.WidthFixed, rowHeaderWidth);
        for (var col = 1; col <= sheet.ColsCount; col++)
        {
            ImGui.TableSetupColumn(ColumnName(col), ImGuiTableColumnFlags.WidthFixed, activeTab.GetColumnWidth(col, defaultColumnWidth));
        }

        ImGui.TableNextRow();
        ImGui.TableSetColumnIndex(0);
        this.DrawHeaderButton("grid-all", "#", new Vector2(rowHeaderWidth - 0.05f, defaultCellHeight - 0.05f), () => this.SelectAllCells(), false);
        for (var col = 1; col <= sheet.ColsCount; col++)
        {
            ImGui.TableSetColumnIndex(col);
            var targetCol = col;
            var columnWidth = MathF.Max(24f, activeTab.GetColumnWidth(col, defaultColumnWidth));
            this.DrawHeaderButton($"col-{col}", ColumnName(col), new Vector2(columnWidth - 0.05f, defaultCellHeight - 0.5f), () => this.SelectColumn(targetCol, ImGui.GetIO().KeyCtrl), false);
            this.DrawColumnResizeHandle(activeTab, col, columnWidth, defaultCellHeight);
        }

        for (var row = 1; row <= sheet.RowsCount; row++)
        {
            var rowHeight = MathF.Max(20f, activeTab.GetRowHeight(row, defaultCellHeight));
            ImGui.TableNextRow();
            ImGui.TableSetColumnIndex(0);
            var targetRow = row;
            this.DrawHeaderButton($"row-{row}", row.ToString(CultureInfo.InvariantCulture), new Vector2(rowHeaderWidth - 0.05f, MathF.Max(1f, rowHeight - 0.05f)), () => this.SelectRow(targetRow, ImGui.GetIO().KeyCtrl), false);
            this.DrawRowResizeHandle(activeTab, row, rowHeight, rowHeaderWidth - 0.05f);

            for (var col = 1; col <= sheet.ColsCount; col++)
            {
                ImGui.TableSetColumnIndex(col);
                var columnWidth = MathF.Max(24f, activeTab.GetColumnWidth(col, defaultColumnWidth));
                this.DrawDataCell(activeTab, row, col, columnWidth, rowHeight, sheet.RowsCount, sheet.ColsCount);
            }
        }

        ImGui.EndTable();

        var overlayClipMin = this.gridClipMin;
        var overlayClipMax = this.gridClipMax;
        this.DrawPendingMergedEditOverlay(activeTab);

        foreach (var overlay in this.pendingMergedTextOverlays)
        {
            var clipMin = new Vector2(MathF.Max(overlay.Min.X, overlayClipMin.X), MathF.Max(overlay.Min.Y, overlayClipMin.Y));
            var clipMax = new Vector2(MathF.Min(overlay.Max.X, overlayClipMax.X), MathF.Min(overlay.Max.Y, overlayClipMax.Y));
            if (clipMax.X > clipMin.X && clipMax.Y > clipMin.Y)
            {
                overlay.DrawList.PushClipRect(clipMin, clipMax, true);
                DrawAlignedCellText(overlay.DrawList, overlay.Min, overlay.Max, overlay.Cell, overlay.Text);
                overlay.DrawList.PopClipRect();
            }
        }

        foreach (var overlay in this.pendingMergedSelectionOverlays)
        {
            var clipMin = new Vector2(MathF.Max(overlay.Min.X, overlayClipMin.X), MathF.Max(overlay.Min.Y, overlayClipMin.Y));
            var clipMax = new Vector2(MathF.Min(overlay.Max.X, overlayClipMax.X), MathF.Min(overlay.Max.Y, overlayClipMax.Y));
            if (clipMax.X > clipMin.X && clipMax.Y > clipMin.Y)
            {
                overlay.DrawList.PushClipRect(clipMin, clipMax, true);
                overlay.DrawList.AddRectFilled(overlay.Min, overlay.Max, ImGui.ColorConvertFloat4ToU32(new Vector4(1f, 0.57f, 0f, 0.08f)), 0f);
                overlay.DrawList.AddRect(overlay.Min, overlay.Max, ImGui.ColorConvertFloat4ToU32(AccentOrange), 0f, 0, 2f);
                overlay.DrawList.PopClipRect();
            }
        }

        ImGui.PopStyleVar();
        ImGui.EndChild();
    }

    private void DrawColumnResizeHandle(SheetTabData activeTab, int col, float columnWidth, float height)
    {
        var itemMin = ImGui.GetItemRectMin();
        var itemMax = ImGui.GetItemRectMax();
        var handleWidth = 18f;
        ImGui.SetCursorScreenPos(new Vector2(itemMax.X - (handleWidth * 0.65f), itemMin.Y));
        ImGui.InvisibleButton($"##col-resize-{col}", new Vector2(handleWidth, height));
        var hovered = ImGui.IsItemHovered();
        var active = ImGui.IsItemActive();

        if (active && Math.Abs(ImGui.GetIO().MouseDelta.X) > 0f)
        {
            activeTab.ColumnWidths[col] = Math.Clamp(columnWidth + ImGui.GetIO().MouseDelta.X, 24f, 420f);
        }

        if (hovered || active)
        {
            var drawList = ImGui.GetWindowDrawList();
            var lineX = itemMax.X - 1f;
            drawList.AddLine(new Vector2(lineX, itemMin.Y + 2f), new Vector2(lineX, itemMax.Y - 2f), ImGui.ColorConvertFloat4ToU32(AccentOrange), 1.2f);
        }
    }

    private void DrawRowResizeHandle(SheetTabData activeTab, int row, float rowHeight, float width)
    {
        var itemMin = ImGui.GetItemRectMin();
        var itemMax = ImGui.GetItemRectMax();
        var handleHeight = 6f;
        ImGui.SetCursorScreenPos(new Vector2(itemMin.X, itemMax.Y - (handleHeight * 0.5f)));
        ImGui.InvisibleButton($"##row-resize-{row}", new Vector2(width, handleHeight));
        var hovered = ImGui.IsItemHovered();
        var active = ImGui.IsItemActive();
        if (active && Math.Abs(ImGui.GetIO().MouseDelta.Y) > 0f)
        {
            activeTab.RowHeights[row] = Math.Clamp(rowHeight + ImGui.GetIO().MouseDelta.Y, 20f, 160f);
        }

        if (hovered || active)
        {
            var drawList = ImGui.GetWindowDrawList();
            var lineY = itemMax.Y - 1f;
            drawList.AddLine(new Vector2(itemMin.X + 2f, lineY), new Vector2(itemMax.X - 2f, lineY), ImGui.ColorConvertFloat4ToU32(AccentOrange), 1.2f);
        }
    }

    private void DrawHeaderButton(string id, string text, Vector2 size, Action onClick, bool active)
    {
        var actualSize = size;
        if (actualSize.X < 0f)
        {
            actualSize.X = ImGui.GetContentRegionAvail().X;
        }

        if (actualSize.Y <= 0f)
        {
            actualSize.Y = 24f;
        }

        var cursor = ImGui.GetCursorScreenPos();
        var clicked = ImGui.InvisibleButton(id, actualSize);
        var hovered = ImGui.IsItemHovered();
        var pressed = ImGui.IsItemActive();
        var drawList = ImGui.GetWindowDrawList();
        var baseColor = active ? AccentOrange : HexToVector4("#3b3b3b");
        var fill = pressed ? AdjustBrightness(baseColor, 0.90f) : hovered ? AdjustBrightness(baseColor, 1.08f) : baseColor;
        var min = cursor;
        var max = cursor + actualSize;
        drawList.AddRectFilled(min + new Vector2(0f, 1f), max + new Vector2(0f, 1f), ImGui.ColorConvertFloat4ToU32(new Vector4(0f, 0f, 0f, 0.14f)), 0f);
        drawList.AddRectFilled(min, max, ImGui.ColorConvertFloat4ToU32(fill), 0f);
        var lineColor = ImGui.ColorConvertFloat4ToU32(new Vector4(0f, 0f, 0f, 0.18f));
        drawList.AddLine(new Vector2(min.X, max.Y - 1f), new Vector2(max.X, max.Y - 1f), lineColor, 1f);
        drawList.AddLine(new Vector2(max.X - 1f, min.Y), new Vector2(max.X - 1f, max.Y), lineColor, 1f);
        if (hovered || pressed)
        {
            var highlight = pressed ? AdjustBrightness(baseColor, 1.16f) : AdjustBrightness(baseColor, 1.22f);
            highlight.W = pressed ? 0.28f : 0.22f;
            drawList.AddRect(min - Vector2.One, max + Vector2.One, ImGui.ColorConvertFloat4ToU32(highlight), 0f, 0, 1.4f);
        }

        var textSize = ImGui.CalcTextSize(text);
        var textPos = new Vector2(
            min.X + ((actualSize.X - textSize.X) / 2f),
            min.Y + ((actualSize.Y - textSize.Y) / 2f));
        var textU32 = ImGui.ColorConvertFloat4ToU32(WhiteText);
        drawList.AddText(textPos, textU32, text);
        drawList.AddText(textPos + new Vector2(0.6f, 0f), textU32, text);

        if (clicked)
        {
            this.CommitAnyInCellEdit();
            onClick();
        }
    }

    private void DrawDataCell(SheetTabData activeTab, int row, int col, float width, float height, int sheetRowsCount, int sheetColsCount)
    {
        var cellKey = GetCellKey(row, col);
        activeTab.TryGetCell(cellKey, out var existingCell);
        if (existingCell is not null && !string.IsNullOrWhiteSpace(existingCell.MergedInto))
        {
            var rootKey = existingCell.MergedInto!;
            ImGui.InvisibleButton($"##mergedproxy-{cellKey}", new Vector2(width, height));
            var editingMergedRoot = this.CanEditCurrentSheet && string.Equals(this.editingCellKey, rootKey, StringComparison.OrdinalIgnoreCase);
            var hoveredProxy = !editingMergedRoot && ImGui.IsItemHovered();
            var clickedProxy = !editingMergedRoot && ImGui.IsItemClicked(ImGuiMouseButton.Left);
            var doubleClickedProxy = !editingMergedRoot && hoveredProxy && ImGui.IsMouseDoubleClicked(ImGuiMouseButton.Left);
            var proxyMin = ImGui.GetItemRectMin();
            var proxyMax = ImGui.GetItemRectMax();
            var hasBlockingProxyLock = this.TryGetBlockingCellLock(rootKey, out var proxyLock) && proxyLock is not null;

            if (activeTab.TryGetCell(rootKey, out var rootCell2) && rootCell2 is not null && TryParseCellKey(rootKey, out var rootRow, out var rootCol))
            {
                var rootStyle = this.GetEffectiveCellStyle(activeTab, rootKey, rootCell2);
                var proxyDrawList = ImGui.GetWindowDrawList();
                var proxyClipMin = new Vector2(MathF.Max(proxyMin.X, this.gridClipMin.X), MathF.Max(proxyMin.Y, this.gridClipMin.Y));
                var proxyClipMax = new Vector2(MathF.Min(proxyMax.X, this.gridClipMax.X), MathF.Min(proxyMax.Y, this.gridClipMax.Y));
                if (proxyClipMax.X > proxyClipMin.X && proxyClipMax.Y > proxyClipMin.Y)
                {
                    var isSelectedProxy = this.selectedCellKeys.Contains(rootKey);
                    var isTopSegment = row == rootRow;
                    var isBottomSegment = row == rootRow + Math.Max(1, rootStyle.RowSpan) - 1;
                    var isLeftSegment = col == rootCol;
                    var isRightSegment = col == rootCol + Math.Max(1, rootStyle.ColSpan) - 1;
                    proxyDrawList.PushClipRect(proxyClipMin, proxyClipMax, true);
                    var rowBaseProxy = row % 2 == 0 ? HexToVector4("#2d2d2d") : HexToVector4("#383838");
                    proxyDrawList.AddRectFilled(proxyMin, proxyMax, ImGui.ColorConvertFloat4ToU32(rowBaseProxy), 0f);
                    if (rootStyle.BackgroundColor != 0)
                    {
                        proxyDrawList.AddRectFilled(proxyMin, proxyMax, rootStyle.BackgroundColor, 0f);
                    }

                    if (hoveredProxy)
                    {
                        proxyDrawList.AddRectFilled(proxyMin, proxyMax, ImGui.ColorConvertFloat4ToU32(new Vector4(1f, 1f, 1f, 0.035f)), 0f);
                    }

                    if (isSelectedProxy)
                    {
                        proxyDrawList.AddRectFilled(proxyMin, proxyMax, ImGui.ColorConvertFloat4ToU32(new Vector4(1f, 0.57f, 0f, 0.08f)), 0f);
                        var accent = ImGui.ColorConvertFloat4ToU32(AccentOrange);
                        if (isTopSegment)
                            proxyDrawList.AddLine(new Vector2(proxyMin.X, proxyMin.Y + 1f), new Vector2(proxyMax.X, proxyMin.Y + 1f), accent, 2f);
                        if (isBottomSegment)
                            proxyDrawList.AddLine(new Vector2(proxyMin.X, proxyMax.Y - 1f), new Vector2(proxyMax.X, proxyMax.Y - 1f), accent, 2f);
                        if (isLeftSegment)
                            proxyDrawList.AddLine(new Vector2(proxyMin.X + 1f, proxyMin.Y), new Vector2(proxyMin.X + 1f, proxyMax.Y), accent, 2f);
                        if (isRightSegment)
                            proxyDrawList.AddLine(new Vector2(proxyMax.X - 1f, proxyMin.Y), new Vector2(proxyMax.X - 1f, proxyMax.Y), accent, 2f);
                    }
                    else
                    {
                        var gridLineColor = ImGui.ColorConvertFloat4ToU32(new Vector4(0f, 0f, 0f, 0.18f));
                        proxyDrawList.AddLine(new Vector2(proxyMin.X, proxyMax.Y - 1f), new Vector2(proxyMax.X, proxyMax.Y - 1f), gridLineColor, 1f);
                        proxyDrawList.AddLine(new Vector2(proxyMax.X - 1f, proxyMin.Y), new Vector2(proxyMax.X - 1f, proxyMax.Y), gridLineColor, 1f);
                    }

                    if (rootStyle.Borders != CellBorderFlags.None)
                    {
                        var borderColor = rootStyle.BorderColor != 0 ? rootStyle.BorderColor : ImGui.ColorConvertFloat4ToU32(new Vector4(1f, 1f, 1f, 0.90f));
                        borderColor = (borderColor & 0x00FFFFFFu) | 0xE6000000u;
                        const float borderThickness = 3.3f;
                        var half = borderThickness * 0.5f;
                        if (rootStyle.Borders.HasFlag(CellBorderFlags.Left) && isLeftSegment)
                            proxyDrawList.AddLine(new Vector2(proxyMin.X + half, proxyMin.Y), new Vector2(proxyMin.X + half, proxyMax.Y), borderColor, borderThickness);
                        if (rootStyle.Borders.HasFlag(CellBorderFlags.Top) && isTopSegment)
                            proxyDrawList.AddLine(new Vector2(proxyMin.X, proxyMin.Y + half), new Vector2(proxyMax.X, proxyMin.Y + half), borderColor, borderThickness);
                        if (rootStyle.Borders.HasFlag(CellBorderFlags.Right) && isRightSegment)
                            proxyDrawList.AddLine(new Vector2(proxyMax.X - half, proxyMin.Y), new Vector2(proxyMax.X - half, proxyMax.Y), borderColor, borderThickness);
                        if (rootStyle.Borders.HasFlag(CellBorderFlags.Bottom) && isBottomSegment)
                            proxyDrawList.AddLine(new Vector2(proxyMin.X, proxyMax.Y - half), new Vector2(proxyMax.X, proxyMax.Y - half), borderColor, borderThickness);
                    }

                    proxyDrawList.PopClipRect();
                }
            }

            if (hoveredProxy && hasBlockingProxyLock && proxyLock is not null)
            {
                ImGui.SetTooltip($"{proxyLock.UserName} is using this cell..");
            }

            if (clickedProxy)
            {
                this.CommitAnyInCellEdit();
                if (ImGui.GetIO().KeyCtrl)
                {
                    if (!this.selectedCellKeys.Add(rootKey))
                    {
                        this.selectedCellKeys.Remove(rootKey);
                    }
                }
                else
                {
                    this.SelectSingleCell(rootKey);
                }

                this.selectedCellKey = this.GetPrimarySelectedCellKey();
            }

            if (doubleClickedProxy && this.CanEditCurrentSheet)
            {
                if (hasBlockingProxyLock && proxyLock is not null)
                {
                    this.statusMessage = $"{proxyLock.UserName} is using this cell..";
                }
                else if (!this.CanModifyCell(rootKey))
                {
                    this.statusMessage = "That cell is protected.";
                }
                else
                {
                    this.SelectSingleCell(rootKey);
                    if (activeTab.TryGetCell(rootKey, out var rootEditCell) && rootEditCell is not null)
                    {
                        this.TryBeginEditingCell(rootKey, rootEditCell.Value ?? string.Empty);
                    }
                    else
                    {
                        this.TryBeginEditingCell(rootKey, string.Empty);
                    }
                }
            }

            return;
        }

        var cell = this.GetEffectiveCellStyle(activeTab, cellKey, existingCell ?? new SheetCellData());
        var colSpan = Math.Max(1, cell.ColSpan);
        var rowSpan = Math.Max(1, cell.RowSpan);
        var mergedWidth = 0f;
        for (var spanCol = col; spanCol < col + colSpan && spanCol <= sheetColsCount; spanCol++)
        {
            mergedWidth += activeTab.GetColumnWidth(spanCol, 140f);
        }
        var mergedHeight = 0f;
        for (var spanRow = row; spanRow < row + rowSpan && spanRow <= sheetRowsCount; spanRow++)
        {
            mergedHeight += activeTab.GetRowHeight(spanRow, 24f);
        }
        var isSelected = this.selectedCellKeys.Contains(cellKey);
        activeTab.TryGetFeature(cellKey, out var cellFeature);

        if (this.CanEditCurrentSheet && string.Equals(this.editingCellKey, cellKey, StringComparison.OrdinalIgnoreCase))
        {
            if (colSpan > 1 || rowSpan > 1)
            {
                this.pendingMergedEditOverlay = new MergedEditOverlay(Min: ImGui.GetCursorScreenPos(), Max: ImGui.GetCursorScreenPos() + new Vector2(mergedWidth, mergedHeight), CellKey: cellKey);
                return;
            }

            var buffer = this.editingCellBuffer;
            ImGui.SetNextItemWidth(mergedWidth);
            var submitted = ImGui.InputText($"##edit-{cellKey}", ref buffer, 4096, ImGuiInputTextFlags.EnterReturnsTrue | ImGuiInputTextFlags.AutoSelectAll);
            this.editingCellBuffer = buffer;
            if (!this.editingFocusRequested)
            {
                ImGui.SetKeyboardFocusHere(-1);
                this.editingFocusRequested = true;
            }

            var deactivatedAfterInitialFrame = ImGui.IsItemDeactivated() && ImGui.GetFrameCount() > this.editingStartedFrame + 1;
            if (submitted || deactivatedAfterInitialFrame)
            {
                this.CommitEditingCell(activeTab, cellKey);
            }

            return;
        }

        ImGui.InvisibleButton($"##cell-{cellKey}", new Vector2(mergedWidth, mergedHeight));
        var hovered = ImGui.IsItemHovered();
        var clicked = ImGui.IsItemClicked(ImGuiMouseButton.Left);
        var doubleClicked = hovered && ImGui.IsMouseDoubleClicked(ImGuiMouseButton.Left);
        var min = ImGui.GetItemRectMin();
        var max = ImGui.GetItemRectMax();
        var drawList = ImGui.GetWindowDrawList();
        var hasBlockingLock = this.TryGetBlockingCellLock(cellKey, out var blockingLock) && blockingLock is not null;

        var visualMax = new Vector2(min.X + mergedWidth, min.Y + mergedHeight);
        var clipMin = new Vector2(MathF.Max(min.X, this.gridClipMin.X), MathF.Max(min.Y, this.gridClipMin.Y));
        var clipMax = new Vector2(MathF.Min(visualMax.X, this.gridClipMax.X), MathF.Min(visualMax.Y, this.gridClipMax.Y));
        if (clipMax.X <= clipMin.X || clipMax.Y <= clipMin.Y)
        {
            return;
        }

        ImGui.PushClipRect(clipMin, clipMax, true);
        var rowBase = row % 2 == 0 ? HexToVector4("#2d2d2d") : HexToVector4("#383838");
        drawList.AddRectFilled(min, visualMax, ImGui.ColorConvertFloat4ToU32(rowBase), 0f);
        if (cell.BackgroundColor != 0)
        {
            drawList.AddRectFilled(min, visualMax, cell.BackgroundColor, 0f);
        }

        if (hovered)
        {
            drawList.AddRectFilled(min, visualMax, ImGui.ColorConvertFloat4ToU32(new Vector4(1f, 1f, 1f, 0.035f)), 0f);
            if (hasBlockingLock && blockingLock is not null)
            {
                ImGui.SetTooltip($"{blockingLock.UserName} is using this cell..");
            }
            else if (cellFeature is not null && cellFeature.Comments.Count > 0)
            {
                ImGui.BeginTooltip();
                foreach (var comment in cellFeature.Comments.TakeLast(3))
                {
                    ImGui.TextDisabled($"{comment.AuthorName} • {comment.CreatedAtUtc.LocalDateTime:t}");
                    ImGui.TextWrapped(comment.Message);
                    ImGui.Separator();
                }
                ImGui.EndTooltip();
            }
        }

        if (isSelected)
        {
            drawList.AddRectFilled(min, visualMax, ImGui.ColorConvertFloat4ToU32(new Vector4(1f, 0.57f, 0f, 0.08f)), 0f);
            drawList.AddRect(min, visualMax, ImGui.ColorConvertFloat4ToU32(AccentOrange), 0f, 0, 2f);
            if (colSpan > 1 || rowSpan > 1)
            {
                this.pendingMergedSelectionOverlays.Add(new MergedSelectionOverlay(drawList, min, visualMax));
            }
        }
        else
        {
            var gridLineColor = ImGui.ColorConvertFloat4ToU32(new Vector4(0f, 0f, 0f, 0.18f));
            drawList.AddLine(new Vector2(min.X, visualMax.Y - 1f), new Vector2(visualMax.X, visualMax.Y - 1f), gridLineColor, 1f);
            drawList.AddLine(new Vector2(visualMax.X - 1f, min.Y), new Vector2(visualMax.X - 1f, visualMax.Y), gridLineColor, 1f);
        }

        this.DrawFeatureBorders(drawList, min, visualMax, cell);
        if (cellFeature is not null && cellFeature.Comments.Count > 0)
        {
            drawList.AddTriangleFilled(new Vector2(visualMax.X - 10f, min.Y + 2f), new Vector2(visualMax.X - 2f, min.Y + 2f), new Vector2(visualMax.X - 2f, min.Y + 10f), ImGui.ColorConvertFloat4ToU32(AccentOrange));
        }
        if (cellFeature is not null && cellFeature.Dropdown is not null)
        {
            drawList.AddTriangleFilled(new Vector2(visualMax.X - 8f, visualMax.Y - 5f), new Vector2(visualMax.X - 3f, visualMax.Y - 5f), new Vector2(visualMax.X - 5.5f, visualMax.Y - 2f), ImGui.ColorConvertFloat4ToU32(WhiteText));
        }

        if (clicked)
        {
            this.CommitAnyInCellEdit();
            if (cellFeature is not null && cellFeature.IsChecklistCell && !ImGui.GetIO().KeyCtrl && !doubleClicked && this.CanModifyCell(cellKey))
            {
                this.PushUndoSnapshot();
                cellFeature.IsChecked = !cellFeature.IsChecked;
                this.RecordActivity("Checklist toggled", cellKey, null, cellFeature.IsChecked ? "checked" : "unchecked");
            }

            if (ImGui.GetIO().KeyCtrl)
            {
                if (!this.selectedCellKeys.Add(cellKey))
                {
                    this.selectedCellKeys.Remove(cellKey);
                }
            }
            else
            {
                this.SelectSingleCell(cellKey);
            }

            this.selectedCellKey = this.GetPrimarySelectedCellKey();
        }

        if (doubleClicked && this.CanEditCurrentSheet)
        {
            if (hasBlockingLock && blockingLock is not null)
            {
                this.statusMessage = $"{blockingLock.UserName} is using this cell..";
            }
            else if (!this.CanModifyCell(cellKey))
            {
                this.statusMessage = "That cell is protected.";
            }
            else
            {
                this.SelectSingleCell(cellKey);
                this.TryBeginEditingCell(cellKey, existingCell?.Value ?? string.Empty);
                if (colSpan > 1 || rowSpan > 1)
                {
                    this.pendingMergedEditOverlay = new MergedEditOverlay(min, visualMax, cellKey);
                    ImGui.PopClipRect();
                    return;
                }
            }
        }

        var displayText = this.GetRenderedCellText(activeTab, cellKey);
        if (colSpan > 1 || rowSpan > 1)
        {
            this.pendingMergedTextOverlays.Add(new MergedTextOverlay(drawList, min, visualMax, cell, displayText));
        }
        else
        {
            DrawAlignedCellText(drawList, min, visualMax, cell, displayText);
        }
        ImGui.PopClipRect();
    }


    private void DrawPendingMergedEditOverlay(SheetTabData activeTab)
    {
        if (this.pendingMergedEditOverlay is null || string.IsNullOrWhiteSpace(this.editingCellKey))
        {
            return;
        }

        var overlay = this.pendingMergedEditOverlay;
        if (!string.Equals(overlay.CellKey, this.editingCellKey, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var clipMin = new Vector2(MathF.Max(overlay.Min.X, this.gridClipMin.X), MathF.Max(overlay.Min.Y, this.gridClipMin.Y));
        var clipMax = new Vector2(MathF.Min(overlay.Max.X, this.gridClipMax.X), MathF.Min(overlay.Max.Y, this.gridClipMax.Y));
        if (clipMax.X <= clipMin.X || clipMax.Y <= clipMin.Y)
        {
            return;
        }

        ImGui.PushClipRect(clipMin, clipMax, true);
        ImGui.SetCursorScreenPos(overlay.Min);
        var buffer = this.editingCellBuffer;
        ImGui.SetNextItemWidth(overlay.Max.X - overlay.Min.X);
        var submitted = ImGui.InputText($"##merged-edit-{overlay.CellKey}", ref buffer, 4096, ImGuiInputTextFlags.EnterReturnsTrue | ImGuiInputTextFlags.AutoSelectAll);
        this.editingCellBuffer = buffer;
        if (!this.editingFocusRequested)
        {
            ImGui.SetKeyboardFocusHere(-1);
            this.editingFocusRequested = true;
        }

        var deactivatedAfterInitialFrame = ImGui.IsItemDeactivated() && ImGui.GetFrameCount() > this.editingStartedFrame + 1;
        if (submitted || deactivatedAfterInitialFrame)
        {
            this.CommitEditingCell(activeTab, overlay.CellKey);
        }

        ImGui.PopClipRect();
    }

    private static void DrawAlignedCellText(ImDrawListPtr drawList, Vector2 min, Vector2 max, SheetCellData cell, string text)
    {
        var normalized = text.Replace("\n", string.Empty);
        var padding = 6f;
        var fontScale = Math.Clamp(cell.FontScale <= 0f ? 1f : cell.FontScale, 0.7f, 2.5f);
        var availableWidth = MathF.Max(0f, (max.X - min.X) - (padding * 2f));
        var font = ImGui.GetFont();
        var fontSize = font.FontSize * fontScale;

        string rendered;
        float textWidth;
        float textHeight;
        if (cell.WrapText)
        {
            rendered = WrapTextToWidth(normalized, availableWidth, fontScale);
            var lines = rendered.Split('\n');
            textWidth = lines.Select(line => ImGui.CalcTextSize(line).X * fontScale).DefaultIfEmpty(0f).Max();
            textHeight = lines.Length * fontSize;
        }
        else
        {
            rendered = FitTextToWidth(normalized.Replace("\n", " "), availableWidth, fontScale);
            var size = ImGui.CalcTextSize(rendered) * fontScale;
            textWidth = size.X;
            textHeight = size.Y;
        }

        var x = min.X + padding;
        if (string.Equals(cell.HorizontalAlign, "center", StringComparison.OrdinalIgnoreCase))
        {
            x = min.X + ((max.X - min.X) - textWidth) / 2f;
        }
        else if (string.Equals(cell.HorizontalAlign, "right", StringComparison.OrdinalIgnoreCase))
        {
            x = max.X - textWidth - padding;
        }

        x = MathF.Max(x, min.X + padding);
        var y = cell.VerticalAlign switch
        {
            "top" => min.Y + 3f,
            "bottom" => max.Y - textHeight - 3f,
            _ => min.Y + ((max.Y - min.Y) - textHeight) / 2f,
        };

        if (cell.WrapText)
        {
            var lineY = y;
            foreach (var line in rendered.Split('\n'))
            {
                drawList.AddText(font, fontSize, new Vector2(x, lineY), cell.TextColor, line);
                if (cell.Bold)
                {
                    drawList.AddText(font, fontSize, new Vector2(x + 0.55f, lineY), cell.TextColor, line);
                }
                if (cell.Italic)
                {
                    drawList.AddText(font, fontSize, new Vector2(x + 0.85f, lineY), cell.TextColor, line);
                }
                if (cell.Underline)
                {
                    var lineWidth = ImGui.CalcTextSize(line).X * fontScale;
                    var underlineY = lineY + fontSize + 1f;
                    drawList.AddLine(new Vector2(x, underlineY), new Vector2(x + lineWidth, underlineY), cell.TextColor, 1f);
                }

                lineY += fontSize;
            }
        }
        else
        {
            drawList.AddText(font, fontSize, new Vector2(x, y), cell.TextColor, rendered);
            if (cell.Bold)
            {
                drawList.AddText(font, fontSize, new Vector2(x + 0.55f, y), cell.TextColor, rendered);
            }
            if (cell.Italic)
            {
                drawList.AddText(font, fontSize, new Vector2(x + 0.85f, y), cell.TextColor, rendered);
            }
            if (cell.Underline)
            {
                var lineY = y + textHeight + 1f;
                drawList.AddLine(new Vector2(x, lineY), new Vector2(x + textWidth, lineY), cell.TextColor, 1f);
            }
        }
    }

    private static string WrapTextToWidth(string text, float maxWidth, float fontScale)
    {
        if (string.IsNullOrWhiteSpace(text) || maxWidth <= 8f)
        {
            return text;
        }

        var words = text.Replace("\n", string.Empty).Split(new[] { ' ', '\n' }, StringSplitOptions.None);
        var lines = new List<string>();
        var current = string.Empty;
        foreach (var rawWord in words)
        {
            var word = rawWord ?? string.Empty;
            var candidate = string.IsNullOrEmpty(current) ? word : current + " " + word;
            if (ImGui.CalcTextSize(candidate).X * fontScale <= maxWidth)
            {
                current = candidate;
                continue;
            }

            if (!string.IsNullOrEmpty(current))
            {
                lines.Add(current);
            }

            current = word;
            while (!string.IsNullOrEmpty(current) && ImGui.CalcTextSize(current).X * fontScale > maxWidth)
            {
                var fitted = FitTextToWidth(current, maxWidth, fontScale);
                if (string.IsNullOrWhiteSpace(fitted) || fitted == "...")
                {
                    break;
                }

                if (fitted.EndsWith("...", StringComparison.Ordinal))
                {
                    fitted = fitted[..^3];
                }

                lines.Add(fitted);
                current = current[fitted.Length..];
            }
        }

        if (!string.IsNullOrEmpty(current))
        {
            lines.Add(current);
        }

        return string.Join("\n", lines);
    }

    private static string FitTextToWidth(string text, float maxWidth, float fontScale)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        if (ImGui.CalcTextSize(text).X * fontScale <= maxWidth)
        {
            return text;
        }

        const string ellipsis = "...";
        var low = 0;
        var high = text.Length;
        var best = ellipsis;
        while (low <= high)
        {
            var mid = (low + high) / 2;
            var candidate = mid >= text.Length ? text : text[..Math.Max(0, mid)] + ellipsis;
            var width = ImGui.CalcTextSize(candidate).X * fontScale;
            if (width <= maxWidth)
            {
                best = candidate;
                low = mid + 1;
            }
            else
            {
                high = mid - 1;
            }
        }

        return best;
    }

    private void DrawDeletePopup()
    {
    }

    private void DrawLocalSavePopup()
    {
        if (!this.localSavePopupOpen)
        {
            return;
        }

        if (this.localSavePopupRequested)
        {
            ImGui.OpenPopup("Local Save Complete");
            this.localSavePopupRequested = false;
        }

        if (ImGui.BeginPopupModal("Local Save Complete", ref this.localSavePopupOpen, ImGuiWindowFlags.AlwaysAutoResize))
        {
            ImGui.TextWrapped("The local file was saved here:");
            ImGui.Spacing();
            ImGui.PushStyleColor(ImGuiCol.Text, AccentBlue);
            if (ImGui.Selectable(this.lastSavedLocalPath, false))
            {
                OpenInExplorer(this.lastSavedLocalPath);
            }
            ImGui.PopStyleColor();
            ImGui.Spacing();
            if (this.DrawStyledButton("close-local-save", "Close", AccentGreen, WhiteText, new Vector2(96, 24)))
            {
                this.localSavePopupOpen = false;
                ImGui.CloseCurrentPopup();
            }

            ImGui.EndPopup();
        }
    }

    private bool CanEditCurrentSheet
        => this.currentSheet is not null
           && this.currentSheet.Data.Settings.ViewMode != SheetViewMode.ReadOnly
           && (this.currentRole == SheetAccessRole.Owner
               || ((this.currentRole == SheetAccessRole.Editor || this.currentRole == SheetAccessRole.Viewer)
                   && this.HasSheetPermission(SheetPermissionType.EditSheet)));

    private bool IsSheetFavorited(string? sheetId)
    {
        if (string.IsNullOrWhiteSpace(sheetId))
        {
            return false;
        }

        this.configuration.SheetFavoriteStates ??= new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        if (this.currentSheet is not null
            && string.Equals(this.currentSheet.Id, sheetId, StringComparison.OrdinalIgnoreCase))
        {
            return this.GetSheetSettings().Favorite;
        }

        if (this.configuration.SheetFavoriteStates.TryGetValue(sheetId, out var persistedFavorite))
        {
            return persistedFavorite;
        }

        if (this.configuration.LocalSheets.TryGetValue(sheetId, out var cachedSheet)
            && cachedSheet?.Data?.Settings is not null)
        {
            var cachedFavorite = cachedSheet.Data.Settings.Favorite;
            this.configuration.SheetFavoriteStates[sheetId] = cachedFavorite;
            return cachedFavorite;
        }

        return false;
    }

    private void SetSheetFavoriteState(string? sheetId, bool favorite)
    {
        if (string.IsNullOrWhiteSpace(sheetId))
        {
            return;
        }

        this.configuration.SheetFavoriteStates ??= new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        this.configuration.SheetFavoriteStates[sheetId] = favorite;

        if (this.configuration.LocalSheets.TryGetValue(sheetId, out var cachedSheet)
            && cachedSheet?.Data?.Settings is not null)
        {
            cachedSheet.Data.Settings.Favorite = favorite;
        }

        if (this.currentSheet is not null
            && string.Equals(this.currentSheet.Id, sheetId, StringComparison.OrdinalIgnoreCase))
        {
            this.GetSheetSettings().Favorite = favorite;
        }
    }

    private async Task ReloadSheetListAsync()
    {
        if (!this.supabase.IsAuthenticated)
        {
            lock (this.accessibleSheetsSync)
            {
                this.accessibleSheets.Clear();
            }
            return;
        }

        var sheets = await this.supabase.ListSheetsAsync().ConfigureAwait(false);
        foreach (var sheet in sheets)
        {
            sheet.Title = StripAutoVersionSuffix(sheet.Title);
        }
        lock (this.accessibleSheetsSync)
        {
            this.accessibleSheets.Clear();
            this.accessibleSheets.AddRange(sheets
                .OrderByDescending(sheet => this.IsSheetFavorited(sheet.Id))
                .ThenByDescending(sheet => this.configuration.RecentSheetAccessTicks.TryGetValue(sheet.Id, out var recentTicks) ? recentTicks : 0L)
                .ThenByDescending(sheet => sheet.UpdatedAt)
                .ThenByDescending(sheet => sheet.CreatedAt)
                .ThenByDescending(sheet => sheet.Id, StringComparer.OrdinalIgnoreCase));
        }
    }

    private async Task LoadSheetAsync(string sheetId)
    {
        var remote = await this.supabase.GetSheetAsync(sheetId).ConfigureAwait(false);
        if (remote is null)
        {
            if (this.configuration.LocalSheets.TryGetValue(sheetId, out var local))
            {
                local.Data.Normalize(local.RowsCount, local.ColsCount);
                this.currentSheet = new RemoteSheet
                {
                    Id = local.SheetId,
                    Title = StripAutoVersionSuffix(local.Title),
                    Code = local.Code,
                    RowsCount = local.RowsCount,
                    ColsCount = local.ColsCount,
                    DefaultRole = local.DefaultRole,
                    Version = local.Version,
                    Data = local.Data,
                    OwnerId = this.configuration.UserId,
                    UpdatedAt = local.LastLocalSaveUtc,
                };
                this.currentRole = SheetAccessRole.Owner;
                this.cachedLatestTemporaryInviteCode = this.GetSheetSettings().TemporaryInviteCodes?.OrderByDescending(x => x.CreatedAtUtc).FirstOrDefault()?.Code;
                this.OnSheetLoaded(this.currentSheet);
                this.ClearSelection();
                this.SelectSingleCell("R1C1");
                this.statusMessage = "Loaded the local cached copy because the cloud version was unavailable.";
                return;
            }

            throw new InvalidOperationException("Sheet not found in Supabase.");
        }

        remote.Data.Normalize(remote.RowsCount, remote.ColsCount);
        remote.Title = StripAutoVersionSuffix(remote.Title);
        var persistedFavorite = this.IsSheetFavorited(remote.Id);
        remote.Data.Settings.Favorite = remote.Data.Settings.Favorite || persistedFavorite;
        this.configuration.SheetFavoriteStates ??= new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        this.configuration.SheetFavoriteStates[remote.Id] = remote.Data.Settings.Favorite;
        this.currentSheet = remote;
        this.currentRole = string.Equals(remote.OwnerId, this.configuration.UserId, StringComparison.OrdinalIgnoreCase)
            ? SheetAccessRole.Owner
            : await this.supabase.GetAccessRoleAsync(remote.Id).ConfigureAwait(false);
        this.cachedLatestTemporaryInviteCode = this.GetSheetSettings().TemporaryInviteCodes?.OrderByDescending(x => x.CreatedAtUtc).FirstOrDefault()?.Code;
        this.configuration.LastOpenedSheetId = remote.Id;
        this.configuration.RecentSheetAccessTicks[remote.Id] = DateTimeOffset.UtcNow.UtcDateTime.Ticks;
        this.configuration.Save();
        this.OnSheetLoaded(remote);
        this.SaveCurrentSheetToLocalCache();
        this.ClearSelection();
        this.SelectSingleCell("R1C1");
        this.statusMessage = $"Loaded {remote.Title} from the cloud.";
    }

    private async Task SaveCurrentSheetToCloudAsync(CancellationToken cancellationToken = default)
    {
        if (this.currentSheet is null)
        {
            return;
        }

        if (!this.CanPersistCurrentSheetChanges())
        {
            throw new InvalidOperationException("This sheet is read-only for your account.");
        }

        await this.sheetSaveGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (this.currentSheet is null)
            {
                return;
            }

            var previousSettings = this.currentSheet.Data.Settings;
            var previousUniqueCode = this.currentOwnerUniqueCode;
            var saved = await this.supabase.UpdateSheetAsync(this.currentSheet, cancellationToken).ConfigureAwait(false);
            saved.Data.Normalize(saved.RowsCount, saved.ColsCount);
            this.EnsureAdvancedDefaults(saved.Data);
            if (previousSettings is not null)
            {
                saved.Data.Settings.Presence = previousSettings.Presence ?? new List<SheetPresenceInfo>();
                saved.Data.Settings.ChatMessages = this.MergeRuntimeChatMessages(previousSettings.ChatMessages, saved.Data.Settings.ChatMessages)
                    .OrderBy(x => x.TimestampUtc)
                    .TakeLast(200)
                    .ToList();
            }

            this.currentSheet = saved;
            this.currentOwnerUniqueCode ??= previousUniqueCode;
            this.OnSheetSavedToCloud(saved);
            this.SaveCurrentSheetToLocalCache();
            this.ApplySavedSheetToAccessibleList(saved);
            if (this.suppressNextCloudSaveStatus)
            {
                this.suppressNextCloudSaveStatus = false;
            }
            else
            {
                this.statusMessage = $"Saved {saved.Title} to the cloud.";
            }
        }
        finally
        {
            this.sheetSaveGate.Release();
        }
    }

    private void ApplySavedSheetToAccessibleList(RemoteSheet saved)
    {
        lock (this.accessibleSheetsSync)
        {
            var existing = this.accessibleSheets.FirstOrDefault(x => string.Equals(x.Id, saved.Id, StringComparison.OrdinalIgnoreCase));
            if (existing is null)
            {
                this.accessibleSheets.Add(new SheetSummary
                {
                    Id = saved.Id,
                    OwnerId = saved.OwnerId,
                    Title = StripAutoVersionSuffix(saved.Title),
                    Code = saved.Code,
                    RowsCount = saved.RowsCount,
                    ColsCount = saved.ColsCount,
                    DefaultRole = saved.DefaultRole,
                    Version = saved.Version,
                    CreatedAt = saved.CreatedAt,
                    UpdatedAt = saved.UpdatedAt,
                });
            }
            else
            {
                existing.Title = StripAutoVersionSuffix(saved.Title);
                existing.Code = saved.Code;
                existing.RowsCount = saved.RowsCount;
                existing.ColsCount = saved.ColsCount;
                existing.DefaultRole = saved.DefaultRole;
                existing.Version = saved.Version;
                existing.UpdatedAt = saved.UpdatedAt;
            }

            this.accessibleSheets.Sort((left, right) =>
            {
                var leftFavorite = this.IsSheetFavorited(left.Id);
                var rightFavorite = this.IsSheetFavorited(right.Id);
                var favoriteCompare = rightFavorite.CompareTo(leftFavorite);
                if (favoriteCompare != 0)
                {
                    return favoriteCompare;
                }

                var leftRecent = this.configuration.RecentSheetAccessTicks.TryGetValue(left.Id, out var leftTicks) ? leftTicks : 0L;
                var rightRecent = this.configuration.RecentSheetAccessTicks.TryGetValue(right.Id, out var rightTicks) ? rightTicks : 0L;
                var recentCompare = rightRecent.CompareTo(leftRecent);
                if (recentCompare != 0)
                {
                    return recentCompare;
                }

                var updatedCompare = right.UpdatedAt.CompareTo(left.UpdatedAt);
                if (updatedCompare != 0)
                {
                    return updatedCompare;
                }

                var createdCompare = right.CreatedAt.CompareTo(left.CreatedAt);
                if (createdCompare != 0)
                {
                    return createdCompare;
                }

                return string.Compare(right.Id, left.Id, StringComparison.OrdinalIgnoreCase);
            });
        }
    }

    private void SaveCurrentSheetToLocalCache()
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var clone = new LocalSheetCache
        {
            SheetId = this.currentSheet.Id,
            Title = StripAutoVersionSuffix(this.currentSheet.Title),
            Code = this.currentSheet.Code,
            RowsCount = this.currentSheet.RowsCount,
            ColsCount = this.currentSheet.ColsCount,
            DefaultRole = this.currentSheet.DefaultRole,
            Version = this.currentSheet.Version,
            Data = SheetSerializationHelper.CloneForLocalCache(this.currentSheet.Data, this.currentSheet.RowsCount, this.currentSheet.ColsCount),
            LastLocalSaveUtc = DateTime.UtcNow,
        };

        clone.Data.Normalize(clone.RowsCount, clone.ColsCount);
        this.configuration.SheetFavoriteStates ??= new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        this.configuration.SheetFavoriteStates[this.currentSheet.Id] = clone.Data.Settings.Favorite;
        this.configuration.LocalSheets[this.currentSheet.Id] = clone;
        this.configuration.LastOpenedSheetId = this.currentSheet.Id;
        this.configuration.Save();
    }

    private string ExportCurrentSheetToLocalFile()
    {
        if (this.currentSheet is null)
        {
            throw new InvalidOperationException("There is no sheet loaded to export.");
        }

        var exportDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "EZSheetsExports");
        Directory.CreateDirectory(exportDirectory);
        var safeTitle = SanitizeFileName(this.currentSheet.Title);
        var safeCode = SanitizeFileName(this.currentSheet.Code);
        var path = Path.Combine(exportDirectory, $"{safeTitle}_{safeCode}.xlsx");

        using var workbook = new XLWorkbook();
        var sheet = this.currentSheet;
        this.currentSheet.Data.Normalize(this.currentSheet.RowsCount, this.currentSheet.ColsCount);
        foreach (var tab in this.currentSheet.Data.Tabs)
        {
            var worksheet = workbook.Worksheets.Add(SanitizeWorksheetName(tab.Name));
            for (var col = 1; col <= sheet.ColsCount; col++)
            {
                worksheet.Column(col).Width = PluginWidthToExcelWidth(tab.GetColumnWidth(col, 140f));
            }
            for (var row = 1; row <= sheet.RowsCount; row++)
            {
                worksheet.Row(row).Height = PluginHeightToExcelHeight(tab.GetRowHeight(row, 24f));
                for (var col = 1; col <= sheet.ColsCount; col++)
                {
                    var key = GetCellKey(row, col);
                    if (!tab.TryGetCell(key, out var cell) || cell is null || !string.IsNullOrWhiteSpace(cell.MergedInto))
                    {
                        continue;
                    }

                    var target = worksheet.Cell(row, col);
                    if (!string.IsNullOrWhiteSpace(cell.Value) && cell.Value.TrimStart().StartsWith("=", StringComparison.Ordinal))
                    {
                        target.FormulaA1 = cell.Value.Trim()[1..];
                    }
                    else
                    {
                        target.Value = cell.Value;
                    }

                    target.Style.Font.Bold = cell.Bold;
                    target.Style.Font.Italic = cell.Italic;
                    target.Style.Font.Underline = cell.Underline ? XLFontUnderlineValues.Single : XLFontUnderlineValues.None;
                    target.Style.Alignment.Horizontal = cell.HorizontalAlign switch
                    {
                        "center" => XLAlignmentHorizontalValues.Center,
                        "right" => XLAlignmentHorizontalValues.Right,
                        _ => XLAlignmentHorizontalValues.Left,
                    };
                    var textColor = ImGuiColorToDrawingColor(cell.TextColor);
                    target.Style.Font.FontColor = XLColor.FromColor(textColor);
                    var effectiveBackgroundColor = cell.BackgroundColor == 0 ? 0xFF2B2B2Bu : cell.BackgroundColor;
                    var fillColor = ImGuiColorToDrawingColor(effectiveBackgroundColor);
                    target.Style.Fill.PatternType = XLFillPatternValues.Solid;
                    target.Style.Fill.BackgroundColor = XLColor.FromColor(fillColor);
                    target.Style.Fill.PatternColor = XLColor.FromColor(fillColor);
                    target.Style.Fill.SetBackgroundColor(XLColor.FromColor(fillColor));
                    target.Style.Fill.SetPatternColor(XLColor.FromColor(fillColor));
                    if (cell.FontScale > 0f)
                    {
                        target.Style.Font.FontSize = 11d * Math.Clamp(cell.FontScale, 0.7f, 2.5f);
                    }
                    if (cell.RowSpan > 1 || cell.ColSpan > 1)
                    {
                        worksheet.Range(row, col, row + cell.RowSpan - 1, col + cell.ColSpan - 1).Merge();
                    }
                }
            }
        }

        workbook.SaveAs(path);
        this.configuration.LastLocalExportPath = path;
        this.configuration.Save();
        return path;
    }

    private static string SanitizeWorksheetName(string input)
    {
        var invalid = new[] { ':', '\\', '/', '?', '*', '[', ']' };
        var sanitized = new string(input.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
        sanitized = string.IsNullOrWhiteSpace(sanitized) ? "Sheet" : sanitized.Trim();
        return sanitized.Length > 31 ? sanitized[..31] : sanitized;
    }

    private static string SanitizeFileName(string input)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var sanitized = new string(input.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
        return string.IsNullOrWhiteSpace(sanitized) ? "sheet" : sanitized;
    }

    private void ApplyToSelectedCells(Action<SheetCellData> apply, bool allowLockedCellOverride = false)
    {
        var sheet = this.currentSheet;
        var activeTab = this.GetActiveTab();
        if (sheet is null || activeTab is null)
        {
            return;
        }

        var anyChanged = false;
        foreach (var key in this.selectedCellKeys)
        {
            if (!this.CanModifyCell(key, allowLockedCellOverride))
            {
                continue;
            }

            var cell = activeTab.GetOrCreateCell(key);
            var before = JsonSerializer.Serialize(cell);
            apply(cell);
            var after = JsonSerializer.Serialize(cell);
            if (!string.Equals(before, after, StringComparison.Ordinal))
            {
                anyChanged = true;
                this.RecordActivity("Cell format changed", key, null, cell.Value);
            }
        }

        if (anyChanged)
        {
            this.MarkPersistableSheetDirty();
        }
    }

    private readonly record struct SelectedRootInfo(string Key, int StartRow, int StartCol, int EndRow, int EndCol);

    private List<SelectedRootInfo> GetSelectedRootInfos(SheetTabData activeTab)
    {
        var result = new Dictionary<string, SelectedRootInfo>(StringComparer.OrdinalIgnoreCase);
        foreach (var selectedKey in this.selectedCellKeys)
        {
            var effectiveKey = selectedKey;
            if (activeTab.TryGetCell(selectedKey, out var selectedCell) && selectedCell is not null && !string.IsNullOrWhiteSpace(selectedCell.MergedInto))
            {
                effectiveKey = selectedCell.MergedInto!;
            }

            if (!TryParseCellKey(effectiveKey, out var row, out var col))
            {
                continue;
            }

            var rootCell = activeTab.GetOrCreateCell(effectiveKey);
            var rowSpan = Math.Max(1, rootCell.RowSpan);
            var colSpan = Math.Max(1, rootCell.ColSpan);
            result[effectiveKey] = new SelectedRootInfo(effectiveKey, row, col, row + rowSpan - 1, col + colSpan - 1);
        }

        return result.Values.ToList();
    }

    private void ApplyBorderPresetToSelection(CellBorderFlags preset, uint color)
    {
        var activeTab = this.GetActiveTab();
        if (activeTab is null)
        {
            return;
        }

        var roots = this.GetSelectedRootInfos(activeTab);
        if (roots.Count == 0)
        {
            return;
        }

        if (preset == CellBorderFlags.None)
        {
            foreach (var info in roots)
            {
                if (!this.CanModifyCell(info.Key))
                {
                    continue;
                }

                var cell = activeTab.GetOrCreateCell(info.Key);
                cell.Borders = CellBorderFlags.None;
                cell.BorderColor = color;
            }

            this.configuration.LastBorderColor = color;
            this.configuration.Save();
            return;
        }

        var applyLeft = preset.HasFlag(CellBorderFlags.Left);
        var applyTop = preset.HasFlag(CellBorderFlags.Top);
        var applyRight = preset.HasFlag(CellBorderFlags.Right);
        var applyBottom = preset.HasFlag(CellBorderFlags.Bottom);
        var minRow = roots.Min(x => x.StartRow);
        var maxRow = roots.Max(x => x.EndRow);
        var minCol = roots.Min(x => x.StartCol);
        var maxCol = roots.Max(x => x.EndCol);

        foreach (var info in roots)
        {
            if (!this.CanModifyCell(info.Key))
            {
                continue;
            }

            var cell = activeTab.GetOrCreateCell(info.Key);
            var next = cell.Borders;
            if (applyLeft)
            {
                next &= ~CellBorderFlags.Left;
                if (info.StartCol == minCol)
                {
                    next |= CellBorderFlags.Left;
                }
            }

            if (applyTop)
            {
                next &= ~CellBorderFlags.Top;
                if (info.StartRow == minRow)
                {
                    next |= CellBorderFlags.Top;
                }
            }

            if (applyRight)
            {
                next &= ~CellBorderFlags.Right;
                if (info.EndCol == maxCol)
                {
                    next |= CellBorderFlags.Right;
                }
            }

            if (applyBottom)
            {
                next &= ~CellBorderFlags.Bottom;
                if (info.EndRow == maxRow)
                {
                    next |= CellBorderFlags.Bottom;
                }
            }

            cell.Borders = next;
            if (next != CellBorderFlags.None)
            {
                cell.BorderColor = color;
            }
        }

        this.configuration.LastBorderColor = color;
        this.configuration.Save();
    }

    private SheetTabData? GetActiveTab()
    {
        if (this.currentSheet is null)
        {
            return null;
        }

        this.currentSheet.Data.Normalize(this.currentSheet.RowsCount, this.currentSheet.ColsCount);
        return this.currentSheet.Data.GetActiveTab(this.currentSheet.RowsCount, this.currentSheet.ColsCount);
    }

    private void SelectSingleCell(string cellKey)
    {
        this.selectedCellKeys.Clear();
        this.selectedCellKeys.Add(cellKey);
        this.selectedCellKey = cellKey;
        this.lastPresenceHeartbeatUtc = DateTimeOffset.MinValue;
    }

    private void SelectRow(int row, bool additive)
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var sheet = this.currentSheet;

        if (!additive)
        {
            this.selectedCellKeys.Clear();
        }

        for (var col = 1; col <= sheet.ColsCount; col++)
        {
            this.selectedCellKeys.Add(GetCellKey(row, col));
        }

        this.selectedCellKey = GetCellKey(row, 1);
        this.lastPresenceHeartbeatUtc = DateTimeOffset.MinValue;
    }

    private void SelectColumn(int col, bool additive)
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var sheet = this.currentSheet;

        if (!additive)
        {
            this.selectedCellKeys.Clear();
        }

        for (var row = 1; row <= sheet.RowsCount; row++)
        {
            this.selectedCellKeys.Add(GetCellKey(row, col));
        }

        this.selectedCellKey = GetCellKey(1, col);
        this.lastPresenceHeartbeatUtc = DateTimeOffset.MinValue;
    }

    private void SelectAllCells()
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var sheet = this.currentSheet;

        this.selectedCellKeys.Clear();
        for (var row = 1; row <= sheet.RowsCount; row++)
        {
            for (var col = 1; col <= sheet.ColsCount; col++)
            {
                this.selectedCellKeys.Add(GetCellKey(row, col));
            }
        }

        this.selectedCellKey = "R1C1";
        this.lastPresenceHeartbeatUtc = DateTimeOffset.MinValue;
    }

    private void ClearSelection()
    {
        this.selectedCellKeys.Clear();
        this.selectedCellKey = "R1C1";
        this.editingCellKey = null;
        this.editingCellBuffer = string.Empty;
        this.editingFocusRequested = false;
        this.editingStartedFrame = 0;
        this.lastPresenceHeartbeatUtc = DateTimeOffset.MinValue;
    }

    private string GetPrimarySelectedCellKey()
    {
        if (this.selectedCellKeys.Contains(this.selectedCellKey))
        {
            return this.selectedCellKey;
        }

        var first = this.selectedCellKeys.FirstOrDefault();
        if (!string.IsNullOrWhiteSpace(first))
        {
            this.selectedCellKey = first;
            return first;
        }

        this.selectedCellKeys.Add("R1C1");
        this.selectedCellKey = "R1C1";
        return this.selectedCellKey;
    }

    private void EnsureSelectionInBounds()
    {
        if (this.currentSheet is null)
        {
            return;
        }

        if (!IsCellWithinSheet(this.selectedCellKey))
        {
            this.SelectSingleCell("R1C1");
        }
    }

    private void BeginEditingCell(string cellKey, string initialValue)
    {
        this.editingCellKey = cellKey;
        this.editingCellBuffer = initialValue;
        this.editingFocusRequested = false;
        this.editingStartedFrame = ImGui.GetFrameCount();
    }

    private void CommitAnyInCellEdit()
    {
        if (string.IsNullOrWhiteSpace(this.editingCellKey))
        {
            return;
        }

        var activeTab = this.GetActiveTab();
        if (activeTab is null)
        {
            this.ReleaseCellLockIfNeeded(this.editingCellKey);
            this.editingCellKey = null;
            this.editingCellBuffer = string.Empty;
            this.editingFocusRequested = false;
            this.editingStartedFrame = 0;
            return;
        }

        this.CommitEditingCell(activeTab, this.editingCellKey);
    }

    private void CommitEditingCell(SheetTabData activeTab, string? cellKey)
    {
        if (string.IsNullOrWhiteSpace(cellKey))
        {
            this.ReleaseCellLockIfNeeded(this.editingCellKey);
            this.editingCellKey = null;
            this.editingCellBuffer = string.Empty;
            this.editingFocusRequested = false;
            this.editingStartedFrame = 0;
            return;
        }

        activeTab.TryGetCell(cellKey, out var existingCell);
        var isMergeRoot = existingCell is not null && (existingCell.RowSpan > 1 || existingCell.ColSpan > 1);
        if (!isMergeRoot)
        {
            isMergeRoot = activeTab.Cells.Values.Any(cell => string.Equals(cell.MergedInto, cellKey, StringComparison.OrdinalIgnoreCase));
        }

        if (!this.CanModifyCell(cellKey))
        {
            this.statusMessage = "That cell is protected.";
            this.ReleaseCellLockIfNeeded(cellKey);
            this.editingCellKey = null;
            this.editingCellBuffer = string.Empty;
            this.editingFocusRequested = false;
            this.editingStartedFrame = 0;
            return;
        }

        var oldValue = existingCell?.Value;
        var value = this.editingCellBuffer;
        var normalizedOldValue = oldValue ?? string.Empty;
        var normalizedNewValue = value ?? string.Empty;
        if (!string.Equals(normalizedOldValue, normalizedNewValue, StringComparison.Ordinal))
        {
            if (string.IsNullOrWhiteSpace(value) && !isMergeRoot)
            {
                activeTab.Cells.Remove(cellKey);
            }
            else
            {
                var cell = existingCell;
                cell ??= activeTab.GetOrCreateCell(cellKey);
                if (cell is null)
                {
                    return;
                }

                cell.Value = value;
                if (cell.RowSpan < 1)
                {
                    cell.RowSpan = 1;
                }

                if (cell.ColSpan < 1)
                {
                    cell.ColSpan = 1;
                }
            }

            this.RecordActivity("Cell edited", cellKey, oldValue, value);
            this.RequestBackgroundCloudSave("Could not sync cell changes.");
        }
        this.SelectSingleCell(cellKey);
        this.ReleaseCellLockIfNeeded(cellKey);
        this.editingCellKey = null;
        this.editingCellBuffer = string.Empty;
        this.editingFocusRequested = false;
        this.editingStartedFrame = 0;
    }

    private string GetDisplayText(SheetTabData tab, string cellKey)
    {
        if (!tab.TryGetCell(cellKey, out var cell) || cell is null || string.IsNullOrWhiteSpace(cell.Value))
        {
            return string.Empty;
        }

        var raw = cell.Value.Trim();
        if (!raw.StartsWith("=", StringComparison.Ordinal))
        {
            return cell.Value;
        }

        if (this.currentSheet is not null && AdvancedFormulaEngine.TryEvaluate(raw, this.currentSheet.Data, this.currentSheet.Data.ActiveTabIndex, out var advancedResult))
        {
            return advancedResult;
        }

        return "#ERR";
    }

    private static bool ShouldUseAdvancedFormulaEngine(string expression)
    {
        return !string.IsNullOrWhiteSpace(expression) && expression.StartsWith("=", StringComparison.Ordinal);
    }

    private FormulaResult EvaluateFormula(SheetTabData tab, string expression, HashSet<string> stack, string currentCellKey)
    {
        var parser = new FormulaParser(expression, reference => this.ResolveReference(tab, reference, stack, currentCellKey));
        return parser.Parse();
    }

    private FormulaResult ResolveReference(SheetTabData tab, string reference, HashSet<string> stack, string currentCellKey)
    {
        if (!TryParseA1Reference(reference, out var row, out var col))
        {
            return FormulaResult.Fail("#REF!");
        }

        var key = GetCellKey(row, col);
        if (string.Equals(key, currentCellKey, StringComparison.OrdinalIgnoreCase) || stack.Contains(key))
        {
            return FormulaResult.Fail("#CYCLE");
        }

        if (!tab.TryGetCell(key, out var cell) || cell is null || string.IsNullOrWhiteSpace(cell.Value))
        {
            return FormulaResult.Ok(0d);
        }

        var raw = cell.Value.Trim();
        if (raw.StartsWith("=", StringComparison.Ordinal))
        {
            stack.Add(key);
            var nested = this.EvaluateFormula(tab, raw[1..], stack, key);
            stack.Remove(key);
            return nested;
        }

        var numericText = raw.Replace(",", string.Empty, StringComparison.Ordinal).Replace(" ", string.Empty, StringComparison.Ordinal);
        return double.TryParse(numericText, NumberStyles.Any, CultureInfo.InvariantCulture, out var numeric)
            ? FormulaResult.Ok(numeric)
            : FormulaResult.Ok(0d);
    }

    private async Task RunActionAsync(Func<Task> action, string fallbackError)
    {
        Interlocked.Increment(ref this.actionInProgressCount);

        try
        {
            await action().ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            this.statusMessage = string.IsNullOrWhiteSpace(ex.Message) ? fallbackError : ex.Message;
            Plugin.Log.Error(ex, fallbackError);
        }
        finally
        {
            Interlocked.Decrement(ref this.actionInProgressCount);
        }
    }

    private bool IsCellWithinSheet(string key)
    {
        if (this.currentSheet is null)
        {
            return false;
        }

        if (!TryParseCellKey(key, out var row, out var col))
        {
            return false;
        }

        return row >= 1 && row <= this.currentSheet.RowsCount && col >= 1 && col <= this.currentSheet.ColsCount;
    }

    private uint GetSheetListColor(string sheetId)
    {
        if (this.configuration.SheetListColors.TryGetValue(sheetId, out var color))
        {
            return color;
        }

        var palette = new[]
        {
            "#ff9100", "#00b0ff", "#83db18", "#d96cff", "#ff6b6b", "#4dd0e1",
        };
        var seed = 17;
        foreach (var ch in sheetId)
        {
            seed = (seed * 31) + ch;
        }
        color = ImGui.ColorConvertFloat4ToU32(HexToVector4(palette[Math.Abs(seed) % palette.Length]));
        this.configuration.SheetListColors[sheetId] = color;
        this.configuration.Save();
        return color;
    }

    private void DrawSectionDivider()
    {
        this.DrawSectionSpacer();
        ImGui.Separator();
        this.DrawSectionSpacer();
    }


    private void DrawCompactSectionDivider()
    {
        this.DrawSectionSpacer(2f);
        ImGui.Separator();
        this.DrawSectionSpacer(2f);
    }

    private void DrawSectionSpacer(float top = 6f, float bottom = 0f)
    {
        ImGui.Dummy(new Vector2(0, top));
        if (bottom > 0f)
        {
            ImGui.Dummy(new Vector2(0, bottom));
        }
    }

    private static Vector4 HexToVector4(string hex)
    {
        var cleaned = hex.Trim().TrimStart('#');
        if (cleaned.Length != 6)
        {
            return WhiteText;
        }

        var r = Convert.ToInt32(cleaned[..2], 16) / 255f;
        var g = Convert.ToInt32(cleaned.Substring(2, 2), 16) / 255f;
        var b = Convert.ToInt32(cleaned.Substring(4, 2), 16) / 255f;
        return new Vector4(r, g, b, 1f);
    }

    private static System.Drawing.Color ImGuiColorToDrawingColor(uint color)
    {
        var r = (byte)(color & 0xFF);
        var g = (byte)((color >> 8) & 0xFF);
        var b = (byte)((color >> 16) & 0xFF);
        var a = (byte)((color >> 24) & 0xFF);
        return System.Drawing.Color.FromArgb(a, r, g, b);
    }

    private static uint DrawingColorToImGuiColor(System.Drawing.Color color)
    {
        return (uint)(color.R | (color.G << 8) | (color.B << 16) | (color.A << 24));
    }

    private static System.Drawing.Color? TryResolveDrawingColor(XLColor? color)
    {
        if (color is null)
        {
            return null;
        }

        try
        {
            var resolved = color.Color;
            if (resolved.A == 0 && resolved.R == 0 && resolved.G == 0 && resolved.B == 0)
            {
                return null;
            }

            return resolved;
        }
        catch
        {
            return null;
        }
    }

    private static System.Drawing.Color? TryResolveCellFillColor(IXLCell cell)
    {
        var direct = TryResolveDrawingColor(cell.Style.Fill.BackgroundColor);
        if (direct.HasValue)
        {
            return direct;
        }

        try
        {
            var fillType = cell.Style.Fill.GetType();
            foreach (var propertyName in new[] { "PatternColor", "ForegroundColor" })
            {
                var prop = fillType.GetProperty(propertyName);
                if (prop?.GetValue(cell.Style.Fill) is XLColor xlColor)
                {
                    var resolved = TryResolveDrawingColor(xlColor);
                    if (resolved.HasValue)
                    {
                        return resolved;
                    }
                }
            }
        }
        catch
        {
        }

        return null;
    }

    private static double PluginWidthToExcelWidth(float pluginWidth) => Math.Max(4d, pluginWidth / 12d);

    private static float ExcelWidthToPluginWidth(double excelWidth) => (float)Math.Clamp(excelWidth * 12d, 60d, 260d);

    private static double PluginHeightToExcelHeight(float pluginHeight) => Math.Max(12d, pluginHeight * 0.75d);

    private static float ExcelHeightToPluginHeight(double excelHeight) => (float)Math.Clamp(excelHeight / 0.75d, 20d, 60d);

    private static Vector4 AdjustBrightness(Vector4 color, float factor)
    {
        return new Vector4(
            Math.Clamp(color.X * factor, 0f, 1f),
            Math.Clamp(color.Y * factor, 0f, 1f),
            Math.Clamp(color.Z * factor, 0f, 1f),
            color.W);
    }

    private bool DrawIconButton(string id, Vector2 size, Vector4 baseColor, Vector4 iconColor, Action<ImDrawListPtr, Vector2, Vector2, bool> drawIcon)
    {
        var actualSize = size;
        if (actualSize.X <= 0f)
        {
            actualSize.X = 24f;
        }

        if (actualSize.Y <= 0f)
        {
            actualSize.Y = 24f;
        }

        var alpha = ImGui.GetStyle().Alpha;
        var cursor = ImGui.GetCursorScreenPos();
        var clicked = ImGui.InvisibleButton(id, actualSize);
        var hovered = ImGui.IsItemHovered();
        var active = ImGui.IsItemActive();
        if (hovered)
        {
            ImGui.SetMouseCursor(ImGuiMouseCursor.Hand);
        }

        var drawList = ImGui.GetWindowDrawList();
        var min = cursor;
        var max = cursor + actualSize;
        var fillColor = active ? AdjustBrightness(baseColor, 0.88f) : hovered ? AdjustBrightness(baseColor, 1.10f) : baseColor;
        fillColor.W *= alpha;
        var borderColor = hovered || active ? AdjustBrightness(baseColor, 1.24f) : AdjustBrightness(baseColor, 0.72f);
        borderColor.W *= alpha;

        drawList.AddRectFilled(min + new Vector2(0f, 1f), max + new Vector2(0f, 1f), ImGui.ColorConvertFloat4ToU32(new Vector4(0f, 0f, 0f, 0.16f * alpha)), 4f);
        drawList.AddRectFilled(min, max, ImGui.ColorConvertFloat4ToU32(fillColor), 4f);
        drawList.AddRect(min, max, ImGui.ColorConvertFloat4ToU32(borderColor), 4f, 0, hovered || active ? 1.2f : 1f);

        if (hovered || active)
        {
            var glow = AdjustBrightness(baseColor, active ? 1.20f : 1.30f);
            glow.W *= active ? 0.28f : 0.20f;
            drawList.AddRect(min - Vector2.One, max + Vector2.One, ImGui.ColorConvertFloat4ToU32(glow), 5f, 0, hovered ? 1.5f : 1.3f);
        }

        drawIcon(drawList, min, max, hovered || active);
        return clicked;
    }

    private static void DrawHomeGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, bool hovered)
    {
        var color = ImGui.ColorConvertFloat4ToU32(new Vector4(1f, 1f, 1f, hovered ? 1f : 0.92f));
        var centerX = (min.X + max.X) * 0.5f;
        var roofTop = new Vector2(centerX, min.Y + 5f);
        var leftRoof = new Vector2(min.X + 5f, min.Y + 11f);
        var rightRoof = new Vector2(max.X - 5f, min.Y + 11f);
        drawList.AddLine(roofTop, leftRoof, color, 1.8f);
        drawList.AddLine(roofTop, rightRoof, color, 1.8f);
        drawList.AddLine(leftRoof, new Vector2(min.X + 5f, max.Y - 5f), color, 1.8f);
        drawList.AddLine(rightRoof, new Vector2(max.X - 5f, max.Y - 5f), color, 1.8f);
        drawList.AddLine(new Vector2(min.X + 5f, max.Y - 5f), new Vector2(max.X - 5f, max.Y - 5f), color, 1.8f);
        drawList.AddLine(new Vector2(centerX - 3f, max.Y - 5f), new Vector2(centerX - 3f, max.Y - 10f), color, 1.8f);
        drawList.AddLine(new Vector2(centerX + 3f, max.Y - 5f), new Vector2(centerX + 3f, max.Y - 10f), color, 1.8f);
        drawList.AddLine(new Vector2(centerX - 3f, max.Y - 10f), new Vector2(centerX + 3f, max.Y - 10f), color, 1.8f);
    }


    private static void DrawOverviewGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, bool hovered)
    {
        var color = ImGui.ColorConvertFloat4ToU32(new Vector4(1f, 1f, 1f, hovered ? 1f : 0.92f));
        var outerMin = new Vector2(min.X + 6f, min.Y + 6f);
        var outerMax = new Vector2(max.X - 6f, max.Y - 6f);
        var innerMin = new Vector2(min.X + 9f, min.Y + 4f);
        var innerMax = new Vector2(max.X - 3f, max.Y - 9f);
        drawList.AddRect(outerMin, outerMax, color, 0f, 0, 1.6f);
        drawList.AddRect(innerMin, innerMax, color, 0f, 0, 1.6f);
    }


    private static void DrawPowerGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, bool hovered)
    {
        var color = ImGui.ColorConvertFloat4ToU32(new Vector4(1f, 1f, 1f, hovered ? 1f : 0.94f));
        var center = (min + max) * 0.5f;
        var radius = MathF.Min(max.X - min.X, max.Y - min.Y) * 0.28f;
        drawList.PathArcTo(center, radius, 0.72f * MathF.PI, 2.28f * MathF.PI, 26);
        drawList.PathStroke(color, ImDrawFlags.None, 1.8f);
        drawList.AddLine(new Vector2(center.X, min.Y + 5f), new Vector2(center.X, center.Y + 1f), color, 2.0f);
    }


    private bool DrawStyledButton(string id, string text, Vector4 baseColor, Vector4 textColor, Vector2 size)
    {
        var actualSize = size;
        if (actualSize.X < 0)
        {
            actualSize.X = ImGui.GetContentRegionAvail().X;
        }

        if (actualSize.Y <= 0)
        {
            actualSize.Y = 24f;
        }

        var alpha = ImGui.GetStyle().Alpha;
        var cursor = ImGui.GetCursorScreenPos();
        var clicked = ImGui.InvisibleButton(id, actualSize);
        var hovered = ImGui.IsItemHovered();
        var active = ImGui.IsItemActive();
        if (hovered)
        {
            ImGui.SetMouseCursor(ImGuiMouseCursor.Hand);
        }
        var drawList = ImGui.GetWindowDrawList();
        var color = active ? AdjustBrightness(baseColor, 0.88f) : hovered ? AdjustBrightness(baseColor, 1.12f) : baseColor;
        color.W *= alpha;
        var borderColor = hovered || active ? AdjustBrightness(baseColor, 1.24f) : AdjustBrightness(baseColor, 0.72f);
        borderColor.W *= alpha;
        var adjustedTextColor = textColor;
        adjustedTextColor.W *= alpha;
        var min = cursor;
        var max = cursor + actualSize;
        drawList.AddRectFilled(min + new Vector2(0f, 1f), max + new Vector2(0f, 1f), ImGui.ColorConvertFloat4ToU32(new Vector4(0f, 0f, 0f, 0.16f * alpha)), 4f);
        drawList.AddRectFilled(min, max, ImGui.ColorConvertFloat4ToU32(color), 4f);
        if (hovered || active)
        {
            var glow = AdjustBrightness(baseColor, active ? 1.20f : 1.30f);
            glow.W *= active ? 0.28f : 0.20f;
            drawList.AddRect(min - Vector2.One, max + Vector2.One, ImGui.ColorConvertFloat4ToU32(glow), 5f, 0, hovered ? 1.5f : 1.3f);
            drawList.AddLine(new Vector2(min.X + 2f, min.Y + 1f), new Vector2(max.X - 2f, min.Y + 1f), ImGui.ColorConvertFloat4ToU32(new Vector4(1f, 1f, 1f, hovered ? 0.16f * alpha : 0.10f * alpha)), 1f);
        }
        drawList.AddRect(min, max, ImGui.ColorConvertFloat4ToU32(borderColor), 4f, 0, hovered || active ? 1.2f : 1f);

        var textSize = ImGui.CalcTextSize(text);
        var textPos = new Vector2(
            min.X + ((actualSize.X - textSize.X) / 2f) + ((text == "→") ? 0.35f : 0f),
            min.Y + ((actualSize.Y - textSize.Y) / 2f) + (((text == "→") || (text == "←")) ? -0.15f : 0f));
        var textU32 = ImGui.ColorConvertFloat4ToU32(adjustedTextColor);
        drawList.AddText(textPos, textU32, text);
        drawList.AddText(textPos + new Vector2(0.6f, 0f), textU32, text);
        return clicked;
    }

    private void DrawFakeBoldText(string text, Vector4 color)
    {
        var pos = ImGui.GetCursorScreenPos();
        var size = ImGui.CalcTextSize(text);
        var drawList = ImGui.GetWindowDrawList();
        var colorU32 = ImGui.ColorConvertFloat4ToU32(color);
        drawList.AddText(pos, colorU32, text);
        drawList.AddText(pos + new Vector2(0.7f, 0f), colorU32, text);
        ImGui.Dummy(new Vector2(size.X + 1f, size.Y + 1f));
    }

    private static string GetCellKey(int row, int col) => $"R{row}C{col}";

    private static bool TryParseCellKey(string key, out int row, out int col)
    {
        row = 0;
        col = 0;
        if (string.IsNullOrWhiteSpace(key) || !key.StartsWith('R'))
        {
            return false;
        }

        var cIndex = key.IndexOf('C');
        if (cIndex <= 1)
        {
            return false;
        }

        return int.TryParse(key.AsSpan(1, cIndex - 1), out row)
               && int.TryParse(key.AsSpan(cIndex + 1), out col);
    }

    private static string ColumnName(int index)
    {
        var name = string.Empty;
        while (index > 0)
        {
            index--;
            name = (char)('A' + (index % 26)) + name;
            index /= 26;
        }

        return name;
    }

    private static string ToA1(string key)
    {
        return TryParseCellKey(key, out var row, out var col)
            ? $"{ColumnName(col)}{row}"
            : key;
    }

    private static bool TryParseA1Reference(string input, out int row, out int col)
    {
        row = 0;
        col = 0;
        if (string.IsNullOrWhiteSpace(input))
        {
            return false;
        }

        input = input.Trim().ToUpperInvariant();
        var split = 0;
        while (split < input.Length && char.IsLetter(input[split]))
        {
            split++;
        }

        if (split == 0 || split == input.Length)
        {
            return false;
        }

        for (var i = 0; i < split; i++)
        {
            col = (col * 26) + (input[i] - 'A' + 1);
        }

        return int.TryParse(input.AsSpan(split), NumberStyles.None, CultureInfo.InvariantCulture, out row) && row > 0 && col > 0;
    }

    private static string StripAutoVersionSuffix(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var trimmed = value.Trim();
        if (trimmed.Length >= 4 && trimmed.EndsWith(")", StringComparison.Ordinal))
        {
            var openIndex = trimmed.LastIndexOf("(v", StringComparison.OrdinalIgnoreCase);
            if (openIndex > 0 && openIndex < trimmed.Length - 3)
            {
                var digits = trimmed.Substring(openIndex + 2, trimmed.Length - openIndex - 3);
                if (digits.All(char.IsDigit) && char.IsWhiteSpace(trimmed[openIndex - 1]))
                {
                    return trimmed[..(openIndex - 1)].TrimEnd();
                }
            }
        }

        var vIndex = trimmed.LastIndexOf(" v", StringComparison.OrdinalIgnoreCase);
        if (vIndex > 0 && vIndex < trimmed.Length - 2)
        {
            var digits = trimmed[(vIndex + 2)..];
            if (digits.All(char.IsDigit))
            {
                return trimmed[..vIndex].TrimEnd();
            }
        }

        return trimmed;
    }

    private static string FormatNumber(double value)
    {
        if (Math.Abs(value % 1d) < 0.0000001d)
        {
            return ((long)Math.Round(value)).ToString(CultureInfo.InvariantCulture);
        }

        return value.ToString("0.##########", CultureInfo.InvariantCulture);
    }

    private static void OpenInExplorer(string filePath)
    {
        var psi = new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = $"/select,\"{filePath}\"",
            UseShellExecute = true,
        };
        Process.Start(psi);
    }



    private void PushUndoSnapshot()
    {
        if (this.currentSheet is null)
        {
            return;
        }

        this.redoHistory.Clear();
        this.undoHistory.Push(SheetSerializationHelper.SerializeForSnapshot(this.currentSheet.Data, this.currentSheet.RowsCount, this.currentSheet.ColsCount));
        while (this.undoHistory.Count > 40)
        {
            var trimmed = this.undoHistory.Reverse().Take(40).Reverse().ToArray();
            this.undoHistory.Clear();
            foreach (var item in trimmed)
            {
                this.undoHistory.Push(item);
            }
        }

        this.PushSharedUndoSnapshot();
    }

    private void PushSharedUndoSnapshot()
    {
        if (this.currentSheet is null || this.sharedHistoryOperationInFlight)
        {
            return;
        }

        var settings = this.GetSheetSettings();
        settings.SharedUndoHistory ??= new List<SheetVersionSnapshot>();
        settings.SharedRedoHistory ??= new List<SheetVersionSnapshot>();
        settings.SharedRedoHistory.Clear();
        settings.SharedUndoHistory.Add(new SheetVersionSnapshot
        {
            Label = $"Undo {DateTime.Now:HH:mm:ss}",
            CreatedAtUtc = DateTimeOffset.UtcNow,
            CreatedByUserId = this.configuration.UserId,
            CreatedByName = this.GetSafeCharacterFullName(),
            DocumentJson = this.SerializeDocumentForSharedHistory(),
        });

        if (settings.SharedUndoHistory.Count > 40)
        {
            settings.SharedUndoHistory.RemoveRange(0, settings.SharedUndoHistory.Count - 40);
        }
    }

    private string SerializeDocumentForSharedHistory()
    {
        return this.currentSheet is null
            ? JsonSerializer.Serialize(SheetDocument.CreateDefault())
            : SheetSerializationHelper.SerializeForSnapshot(this.currentSheet.Data, this.currentSheet.RowsCount, this.currentSheet.ColsCount);
    }

    private bool TryUndoSharedChange()
    {
        if (this.currentSheet is null || this.sharedHistoryOperationInFlight)
        {
            return false;
        }

        var settings = this.GetSheetSettings();
        settings.SharedUndoHistory ??= new List<SheetVersionSnapshot>();
        settings.SharedRedoHistory ??= new List<SheetVersionSnapshot>();
        if (settings.SharedUndoHistory.Count == 0)
        {
            return false;
        }

        this.sharedHistoryOperationInFlight = true;
        try
        {
            settings.SharedRedoHistory.Add(new SheetVersionSnapshot
            {
                Label = $"Redo {DateTime.Now:HH:mm:ss}",
                CreatedAtUtc = DateTimeOffset.UtcNow,
                CreatedByUserId = this.configuration.UserId,
                CreatedByName = this.GetSafeCharacterFullName(),
                DocumentJson = this.SerializeDocumentForSharedHistory(),
            });
            if (settings.SharedRedoHistory.Count > 40)
            {
                settings.SharedRedoHistory.RemoveRange(0, settings.SharedRedoHistory.Count - 40);
            }

            var snapshot = settings.SharedUndoHistory[^1];
            settings.SharedUndoHistory.RemoveAt(settings.SharedUndoHistory.Count - 1);
            var restored = JsonSerializer.Deserialize<SheetDocument>(snapshot.DocumentJson) ?? SheetDocument.CreateDefault();
            restored.Normalize(this.currentSheet.RowsCount, this.currentSheet.ColsCount);
            this.currentSheet.Data = restored;
            this.statusMessage = "Undid the last shared change.";
            return true;
        }
        finally
        {
            this.sharedHistoryOperationInFlight = false;
        }
    }

    private bool TryRedoSharedChange()
    {
        if (this.currentSheet is null || this.sharedHistoryOperationInFlight)
        {
            return false;
        }

        var settings = this.GetSheetSettings();
        settings.SharedUndoHistory ??= new List<SheetVersionSnapshot>();
        settings.SharedRedoHistory ??= new List<SheetVersionSnapshot>();
        if (settings.SharedRedoHistory.Count == 0)
        {
            return false;
        }

        this.sharedHistoryOperationInFlight = true;
        try
        {
            settings.SharedUndoHistory.Add(new SheetVersionSnapshot
            {
                Label = $"Undo {DateTime.Now:HH:mm:ss}",
                CreatedAtUtc = DateTimeOffset.UtcNow,
                CreatedByUserId = this.configuration.UserId,
                CreatedByName = this.GetSafeCharacterFullName(),
                DocumentJson = this.SerializeDocumentForSharedHistory(),
            });
            if (settings.SharedUndoHistory.Count > 40)
            {
                settings.SharedUndoHistory.RemoveRange(0, settings.SharedUndoHistory.Count - 40);
            }

            var snapshot = settings.SharedRedoHistory[^1];
            settings.SharedRedoHistory.RemoveAt(settings.SharedRedoHistory.Count - 1);
            var restored = JsonSerializer.Deserialize<SheetDocument>(snapshot.DocumentJson) ?? SheetDocument.CreateDefault();
            restored.Normalize(this.currentSheet.RowsCount, this.currentSheet.ColsCount);
            this.currentSheet.Data = restored;
            this.statusMessage = "Redid the last shared change.";
            return true;
        }
        finally
        {
            this.sharedHistoryOperationInFlight = false;
        }
    }

    private void UndoLastChange()
    {
        if (this.TryUndoSharedChange())
        {
            return;
        }

        if (this.currentSheet is null || this.undoHistory.Count == 0)
        {
            return;
        }

        this.redoHistory.Push(SheetSerializationHelper.SerializeForSnapshot(this.currentSheet.Data, this.currentSheet.RowsCount, this.currentSheet.ColsCount));
        var snapshot = this.undoHistory.Pop();
        var restored = JsonSerializer.Deserialize<SheetDocument>(snapshot) ?? SheetDocument.CreateDefault();
        restored.Normalize(this.currentSheet.RowsCount, this.currentSheet.ColsCount);
        this.currentSheet.Data = restored;
        this.statusMessage = "Undid the last change.";
    }

    private void MergeSelectedCells()
    {
        if (!this.CanEditCurrentSheet || this.currentSheet is null)
        {
            return;
        }

        var activeTab = this.GetActiveTab();
        if (activeTab is null)
        {
            return;
        }

        var positions = this.selectedCellKeys
            .Select(key => TryParseCellKey(key, out var r, out var c) ? (key, row: r, col: c) : (key, row: 0, col: 0))
            .Where(x => x.row > 0 && x.col > 0)
            .ToList();
        if (positions.Count == 0)
        {
            return;
        }

        var minRow = positions.Min(x => x.row);
        var maxRow = positions.Max(x => x.row);
        var minCol = positions.Min(x => x.col);
        var maxCol = positions.Max(x => x.col);
        var rootKey = GetCellKey(minRow, minCol);

        if (positions.Count == 1 && activeTab.TryGetCell(rootKey, out var singleCell) && singleCell is not null && (singleCell.RowSpan > 1 || singleCell.ColSpan > 1))
        {
            this.PushUndoSnapshot();
            this.UnmergeBlock(activeTab, rootKey, singleCell.RowSpan, singleCell.ColSpan);
            this.SelectSingleCell(rootKey);
            this.statusMessage = "Unmerged the selected cells.";
            return;
        }

        var isRectangular = positions.Count == ((maxRow - minRow + 1) * (maxCol - minCol + 1));
        if (!isRectangular)
        {
            this.statusMessage = "Select a rectangular block before merging cells.";
            return;
        }

        if (activeTab.TryGetCell(rootKey, out var existingRoot) && existingRoot is not null)
        {
            var existingRowSpan = Math.Max(1, existingRoot.RowSpan);
            var existingColSpan = Math.Max(1, existingRoot.ColSpan);
            var matchesExistingMerge = (existingRowSpan > 1 || existingColSpan > 1)
                && existingRowSpan == (maxRow - minRow + 1)
                && existingColSpan == (maxCol - minCol + 1)
                && positions.All(pos => pos.key.Equals(rootKey, StringComparison.OrdinalIgnoreCase)
                    || (activeTab.TryGetCell(pos.key, out var child) && child is not null && string.Equals(child.MergedInto, rootKey, StringComparison.OrdinalIgnoreCase)));
            if (matchesExistingMerge)
            {
                this.PushUndoSnapshot();
                this.UnmergeBlock(activeTab, rootKey, existingRowSpan, existingColSpan);
                this.SelectSingleCell(rootKey);
                this.statusMessage = "Unmerged the selected cells.";
                return;
            }
        }

        if (positions.Count < 2)
        {
            return;
        }

        this.PushUndoSnapshot();
        foreach (var pos in positions)
        {
            var cell = activeTab.GetOrCreateCell(pos.key);
            cell.MergedInto = null;
            cell.RowSpan = 1;
            cell.ColSpan = 1;
        }

        var rootCell = activeTab.GetOrCreateCell(rootKey);
        rootCell.MergedInto = null;
        rootCell.RowSpan = maxRow - minRow + 1;
        rootCell.ColSpan = maxCol - minCol + 1;
        foreach (var pos in positions)
        {
            if (pos.key.Equals(rootKey, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var cell = activeTab.GetOrCreateCell(pos.key);
            cell.MergedInto = rootKey;
            cell.RowSpan = 1;
            cell.ColSpan = 1;
        }

        this.SelectSingleCell(rootKey);
        this.statusMessage = "Merged the selected cells.";
    }

    private void UnmergeBlock(SheetTabData activeTab, string rootKey, int rowSpan, int colSpan)
    {
        if (!TryParseCellKey(rootKey, out var rootRow, out var rootCol))
        {
            return;
        }

        var rootCell = activeTab.GetOrCreateCell(rootKey);
        rootCell.MergedInto = null;
        rootCell.RowSpan = 1;
        rootCell.ColSpan = 1;

        for (var row = rootRow; row < rootRow + Math.Max(1, rowSpan); row++)
        {
            for (var col = rootCol; col < rootCol + Math.Max(1, colSpan); col++)
            {
                var key = GetCellKey(row, col);
                if (key.Equals(rootKey, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var cell = activeTab.GetOrCreateCell(key);
                if (string.Equals(cell.MergedInto, rootKey, StringComparison.OrdinalIgnoreCase))
                {
                    cell.MergedInto = null;
                }

                cell.RowSpan = 1;
                cell.ColSpan = 1;
            }
        }
    }

    private void ImportSheetFromFile()
    {
        lock (this.importDialogSync)
        {
            if (this.importDialogOpen)
            {
                return;
            }

            this.importDialogOpen = true;
            this.pendingImportPath = null;
            this.pendingImportError = null;
        }

        var thread = new Thread(() =>
        {
            try
            {
                using var dialog = new WinForms.OpenFileDialog
                {
                    Filter = "Excel Workbook (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv",
                    Multiselect = false,
                    CheckFileExists = true,
                    RestoreDirectory = true,
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    Title = "Import Sheet"
                };

                var dialogResult = dialog.ShowDialog();
                lock (this.importDialogSync)
                {
                    if (dialogResult == WinForms.DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.FileName) && File.Exists(dialog.FileName))
                    {
                        this.pendingImportPath = dialog.FileName;
                    }
                    else
                    {
                        this.pendingImportError = "Import canceled.";
                    }
                }
            }
            catch (Exception ex)
            {
                lock (this.importDialogSync)
                {
                    this.pendingImportError = ex.Message;
                }
            }
            finally
            {
                lock (this.importDialogSync)
                {
                    this.importDialogOpen = false;
                }
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();
        this.statusMessage = "Waiting for file selection...";
    }

    private void ProcessPendingImportRequest()
    {
        string? importPath = null;
        string? importError = null;
        lock (this.importDialogSync)
        {
            if (!string.IsNullOrWhiteSpace(this.pendingImportPath))
            {
                importPath = this.pendingImportPath;
                this.pendingImportPath = null;
            }

            if (!string.IsNullOrWhiteSpace(this.pendingImportError))
            {
                importError = this.pendingImportError;
                this.pendingImportError = null;
            }
        }

        if (!string.IsNullOrWhiteSpace(importError))
        {
            this.statusMessage = importError;
            return;
        }

        if (string.IsNullOrWhiteSpace(importPath))
        {
            return;
        }

        try
        {
            this.PushUndoSnapshot();
            if (importPath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            {
                this.ImportCsv(importPath);
            }
            else
            {
                this.ImportXlsx(importPath);
            }
        }
        catch (Exception ex)
        {
            this.statusMessage = ex.Message;
        }
    }

    private void ImportCsv(string path)
    {
        if (this.currentSheet is null)
        {
            return;
        }

        var lines = File.ReadAllLines(path);
        var sheet = this.currentSheet;
        var activeTab = this.GetActiveTab();
        if (sheet is null || activeTab is null)
        {
            return;
        }

        activeTab.Cells.Clear();
        for (var row = 0; row < Math.Min(lines.Length, this.currentSheet.RowsCount); row++)
        {
            var cols = lines[row].Split(',');
            for (var col = 0; col < Math.Min(cols.Length, this.currentSheet.ColsCount); col++)
            {
                var value = cols[col].Trim();
                if (string.IsNullOrEmpty(value))
                {
                    continue;
                }

                activeTab.GetOrCreateCell(GetCellKey(row + 1, col + 1)).Value = value;
            }
        }

        this.statusMessage = $"Imported {Path.GetFileName(path)}.";
    }

    private void ImportXlsx(string path)
    {
        if (this.currentSheet is null)
        {
            return;
        }

        using var workbook = new XLWorkbook(path);
        var remoteSheet = this.currentSheet;
        var worksheets = workbook.Worksheets.ToList();
        if (worksheets.Count == 0)
        {
            this.statusMessage = "The selected workbook does not contain any worksheets.";
            return;
        }

        const int maxImportRows = 300;
        const int maxImportCols = 80;
        var importedRows = 1;
        var importedCols = 1;

        foreach (var worksheet in worksheets)
        {
            var usedRange = worksheet.RangeUsed();
            if (usedRange is not null)
            {
                importedRows = Math.Max(importedRows, usedRange.RangeAddress.LastAddress.RowNumber);
                importedCols = Math.Max(importedCols, usedRange.RangeAddress.LastAddress.ColumnNumber);
            }

            foreach (var mergedRange in worksheet.MergedRanges)
            {
                importedRows = Math.Max(importedRows, mergedRange.RangeAddress.LastAddress.RowNumber);
                importedCols = Math.Max(importedCols, mergedRange.RangeAddress.LastAddress.ColumnNumber);
            }
        }

        remoteSheet.RowsCount = Math.Clamp(importedRows, 1, maxImportRows);
        remoteSheet.ColsCount = Math.Clamp(importedCols, 1, maxImportCols);

        var document = new SheetDocument { Tabs = new List<SheetTabData>(), ActiveTabIndex = 0 };
        foreach (var worksheet in worksheets)
        {
            var tab = new SheetTabData { Name = worksheet.Name };
            for (var col = 1; col <= remoteSheet.ColsCount; col++)
            {
                var worksheetColumn = worksheet.Column(col);
                var importedWidth = worksheetColumn.IsHidden || worksheetColumn.Width <= 0d
                    ? 140f
                    : ExcelWidthToPluginWidth(worksheetColumn.Width);
                tab.ColumnWidths[col] = Math.Clamp(importedWidth, 60f, 260f);
            }

            for (var row = 1; row <= remoteSheet.RowsCount; row++)
            {
                var worksheetRow = worksheet.Row(row);
                var importedHeight = worksheetRow.IsHidden || worksheetRow.Height <= 0d
                    ? 24f
                    : ExcelHeightToPluginHeight(worksheetRow.Height);
                tab.RowHeights[row] = Math.Clamp(importedHeight, 20f, 60f);
            }

            var mergedRanges = worksheet.MergedRanges.ToList();
            for (var row = 1; row <= remoteSheet.RowsCount; row++)
            {
                for (var col = 1; col <= remoteSheet.ColsCount; col++)
                {
                    var source = worksheet.Cell(row, col);
                    var key = GetCellKey(row, col);
                    var inMerged = mergedRanges.Any(range => row >= range.RangeAddress.FirstAddress.RowNumber && row <= range.RangeAddress.LastAddress.RowNumber && col >= range.RangeAddress.FirstAddress.ColumnNumber && col <= range.RangeAddress.LastAddress.ColumnNumber);
                    if (!IsMeaningfulImportCell(source, inMerged))
                    {
                        continue;
                    }

                    var cell = tab.GetOrCreateCell(key);
                    cell.Value = !string.IsNullOrWhiteSpace(source.FormulaA1) ? "=" + source.FormulaA1 : source.GetFormattedString();
                    cell.Bold = source.Style.Font.Bold;
                    cell.Italic = source.Style.Font.Italic;
                    cell.Underline = source.Style.Font.Underline != XLFontUnderlineValues.None;
                    cell.HorizontalAlign = source.Style.Alignment.Horizontal switch
                    {
                        XLAlignmentHorizontalValues.Center => "center",
                        XLAlignmentHorizontalValues.Right => "right",
                        _ => "left",
                    };
                    cell.FontScale = (float)Math.Clamp(source.Style.Font.FontSize / 11d, 0.7d, 2.5d);
                    var fontColor = TryResolveDrawingColor(source.Style.Font.FontColor);
                    if (fontColor.HasValue)
                    {
                        cell.TextColor = DrawingColorToImGuiColor(fontColor.Value);
                    }

                    var fill = TryResolveCellFillColor(source);
                    if (fill.HasValue)
                    {
                        cell.BackgroundColor = DrawingColorToImGuiColor(fill.Value);
                    }
                }
            }

            foreach (var mergedRange in mergedRanges)
            {
                var first = mergedRange.FirstCell();
                var last = mergedRange.LastCell();
                if (first.Address.RowNumber < 1 || first.Address.ColumnNumber < 1 || first.Address.RowNumber > remoteSheet.RowsCount || first.Address.ColumnNumber > remoteSheet.ColsCount)
                {
                    continue;
                }

                var rootKey = GetCellKey(first.Address.RowNumber, first.Address.ColumnNumber);
                var rootCell = tab.GetOrCreateCell(rootKey);
                rootCell.RowSpan = Math.Max(1, Math.Min(remoteSheet.RowsCount, last.Address.RowNumber) - first.Address.RowNumber + 1);
                rootCell.ColSpan = Math.Max(1, Math.Min(remoteSheet.ColsCount, last.Address.ColumnNumber) - first.Address.ColumnNumber + 1);
                for (var mergeRow = first.Address.RowNumber; mergeRow <= Math.Min(remoteSheet.RowsCount, last.Address.RowNumber); mergeRow++)
                {
                    for (var mergeCol = first.Address.ColumnNumber; mergeCol <= Math.Min(remoteSheet.ColsCount, last.Address.ColumnNumber); mergeCol++)
                    {
                        var key = GetCellKey(mergeRow, mergeCol);
                        if (string.Equals(key, rootKey, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        tab.GetOrCreateCell(key).MergedInto = rootKey;
                    }
                }
            }

            document.Tabs.Add(tab);
        }

        document.Normalize(remoteSheet.RowsCount, remoteSheet.ColsCount);
        remoteSheet.Data = document;
        this.ClearSelection();
        this.SelectSingleCell("R1C1");
        this.statusMessage = $"Imported {Path.GetFileName(path)}.";
    }

    private static bool IsMeaningfulImportCell(IXLCell cell, bool inMergedRange)
    {
        if (!string.IsNullOrWhiteSpace(cell.FormulaA1))
        {
            return true;
        }

        var value = cell.GetFormattedString();
        if (!string.IsNullOrWhiteSpace(value))
        {
            return true;
        }

        if (inMergedRange)
        {
            return true;
        }

        if (cell.Style.Font.Bold || cell.Style.Font.Italic || cell.Style.Font.Underline != XLFontUnderlineValues.None)
        {
            return true;
        }

        if (cell.Style.Alignment.Horizontal == XLAlignmentHorizontalValues.Center || cell.Style.Alignment.Horizontal == XLAlignmentHorizontalValues.Right)
        {
            return true;
        }

        if (TryResolveDrawingColor(cell.Style.Font.FontColor).HasValue)
        {
            return true;
        }

        if (TryResolveCellFillColor(cell).HasValue)
        {
            return true;
        }

        return false;
    }

    private void DrawHelpWindow()
    {
        if (!this.helpWindowOpen)
        {
            return;
        }

        ImGui.SetNextWindowSize(new Vector2(640f, 420f), ImGuiCond.FirstUseEver);
        if (ImGui.Begin("EZSheets Help###EZSheetsHelp", ref this.helpWindowOpen))
        {
            this.DrawFakeBoldText("How to use EZSheets", AccentOrange);
            ImGui.Spacing();
            ImGui.TextWrapped("1. Sign in with Discord. 2. Create a new synced sheet or access one with a shared code. 3. Double-click any cell to edit it. 4. Use the toolbar buttons to style text, fill colors, align content, change text size, or merge a rectangular selection. 5. Save to Cloud when you want everyone with access to receive the latest version. 6. Save Local exports your current sheet to an Excel file on your PC.");
            ImGui.Spacing();
            ImGui.Separator();
            ImGui.Spacing();
            ImGui.TextWrapped("Quick test area:");
            var preview = "A1";
            this.DrawStyledButton("help-preview-cell", preview, HexToVector4("#3b3b3b"), WhiteText, new Vector2(80, 22));
            ImGui.SameLine();
            ImGui.TextWrapped("Use the same toolbar in the main window to test how styling behaves on real cells.");
        }
        ImGui.End();
    }

    private static void DrawUndoGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max)
    {
        var color = ImGui.ColorConvertFloat4ToU32(WhiteText);
        var text = "←";
        var textSize = ImGui.CalcTextSize(text);
        var pos = new Vector2(min.X + ((max.X - min.X) - textSize.X) / 2f, min.Y + ((max.Y - min.Y) - textSize.Y) / 2f - 1f);
        drawList.AddText(pos, color, text);
        drawList.AddText(pos + new Vector2(0.6f, 0f), color, text);
    }

    private void DrawUndoGlyphLastItem()
    {
    }

    private string GetSafeCharacterFullName()
    {
        try
        {
            var localPlayer = Plugin.ObjectTable?.LocalPlayer;
            var rawName = localPlayer?.Name.TextValue?.Trim() ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(rawName))
            {
                if (!string.Equals(this.configuration.LastCharacterFullName, rawName, StringComparison.Ordinal))
                {
                    this.configuration.LastCharacterFullName = rawName;
                    this.configuration.Save();
                }

                return rawName;
            }
        }
        catch
        {
        }

        var cachedName = this.configuration.LastCharacterFullName?.Trim() ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(cachedName))
        {
            return cachedName;
        }

        var displayName = this.configuration.UserDisplayName?.Trim() ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(displayName) && !displayName.Contains('@', StringComparison.Ordinal))
        {
            return displayName.TrimStart('.');
        }

        return "Player";
    }

    private string GetSafeDiscordDisplayName()
    {
        var fullName = this.GetSafeCharacterFullName();
        var first = fullName.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).FirstOrDefault();
        if (!string.IsNullOrWhiteSpace(first))
        {
            return first;
        }

        return "Player";
    }

    private void DrawCompactFormatToolbar(SheetTabData activeTab, string primaryKey, SheetCellData styleSourceCell)
    {
        var settings = this.GetSheetSettings();
        var textColor = ImGui.ColorConvertU32ToFloat4(styleSourceCell.TextColor);
        var backgroundColor = styleSourceCell.BackgroundColor == 0 ? new Vector4(1f, 1f, 1f, 1f) : ImGui.ColorConvertU32ToFloat4(styleSourceCell.BackgroundColor);
        var align = styleSourceCell.HorizontalAlign;
        var buttonSize = new Vector2(26f, 22f);

        if (this.DrawToolButton("fmt-text-color", buttonSize, "Text color", (dl, min, max, hovered) => DrawTextColorGlyph(dl, min, max, textColor), false))
        {
            ImGui.OpenPopup("fmt-text-color-popup");
        }

        if (ImGui.BeginPopup("fmt-text-color-popup"))
        {
            var picker = textColor;
            if (ImGui.ColorPicker4("##textColorPicker", ref picker, ImGuiColorEditFlags.NoSidePreview | ImGuiColorEditFlags.NoSmallPreview | ImGuiColorEditFlags.AlphaBar))
            {
                this.PushUndoSnapshot();
                var applied = ImGui.ColorConvertFloat4ToU32(picker);
                this.ApplyToSelectedCells(cell => cell.TextColor = applied);
                textColor = picker;
            }
            ImGui.EndPopup();
        }

        ImGui.SameLine(0, 4f);
        if (this.DrawToolButton("fmt-fill-color", buttonSize, "Fill color", (dl, min, max, hovered) => DrawFillColorGlyph(dl, min, max, backgroundColor), false))
        {
            ImGui.OpenPopup("fmt-fill-color-popup");
        }

        if (ImGui.BeginPopup("fmt-fill-color-popup"))
        {
            var picker = backgroundColor;
            if (ImGui.ColorPicker4("##fillColorPicker", ref picker, ImGuiColorEditFlags.NoSidePreview | ImGuiColorEditFlags.NoSmallPreview | ImGuiColorEditFlags.NoAlpha))
            {
                this.PushUndoSnapshot();
                var opaque = new Vector4(picker.X, picker.Y, picker.Z, 1f);
                var applied = ImGui.ColorConvertFloat4ToU32(opaque);
                this.ApplyToSelectedCells(cell => cell.BackgroundColor = applied);
                backgroundColor = opaque;
            }
            ImGui.EndPopup();
        }

        ImGui.SameLine(0, 4f);
        if (this.DrawToolButton("fmt-align-left", buttonSize, "Align left", (dl, min, max, hovered) => DrawAlignGlyph(dl, min, max, 0), string.Equals(align, "left", StringComparison.OrdinalIgnoreCase)))
        {
            this.PushUndoSnapshot();
            this.ApplyToSelectedCells(cell => cell.HorizontalAlign = "left");
        }

        ImGui.SameLine(0, 4f);
        if (this.DrawToolButton("fmt-align-center", buttonSize, "Align center", (dl, min, max, hovered) => DrawAlignGlyph(dl, min, max, 1), string.Equals(align, "center", StringComparison.OrdinalIgnoreCase)))
        {
            this.PushUndoSnapshot();
            this.ApplyToSelectedCells(cell => cell.HorizontalAlign = "center");
        }

        ImGui.SameLine(0, 4f);
        if (this.DrawToolButton("fmt-align-right", buttonSize, "Align right", (dl, min, max, hovered) => DrawAlignGlyph(dl, min, max, 2), string.Equals(align, "right", StringComparison.OrdinalIgnoreCase)))
        {
            this.PushUndoSnapshot();
            this.ApplyToSelectedCells(cell => cell.HorizontalAlign = "right");
        }

        ImGui.SameLine(0, 4f);
        if (this.DrawToolButton("fmt-font-up", buttonSize, "Increase text size", DrawFontUpGlyph, false))
        {
            this.PushUndoSnapshot();
            this.ApplyToSelectedCells(cell => cell.FontScale = Math.Clamp((cell.FontScale <= 0f ? 1f : cell.FontScale) + 0.1f, 0.7f, 2.5f));
        }

        ImGui.SameLine(0, 4f);
        if (this.DrawToolButton("fmt-font-down", buttonSize, "Decrease text size", DrawFontDownGlyph, false))
        {
            this.PushUndoSnapshot();
            this.ApplyToSelectedCells(cell => cell.FontScale = Math.Clamp((cell.FontScale <= 0f ? 1f : cell.FontScale) - 0.1f, 0.7f, 2.5f));
        }

        ImGui.SameLine(0, 4f);
        if (this.DrawToolButton("fmt-bold", buttonSize, "Bold", DrawBoldGlyph, styleSourceCell.Bold))
        {
            this.PushUndoSnapshot();
            var newValue = !styleSourceCell.Bold;
            this.ApplyToSelectedCells(cell => cell.Bold = newValue);
        }

        ImGui.SameLine(0, 4f);
        if (this.DrawToolButton("fmt-italic", buttonSize, "Italic", DrawItalicGlyph, styleSourceCell.Italic))
        {
            this.PushUndoSnapshot();
            var newValue = !styleSourceCell.Italic;
            this.ApplyToSelectedCells(cell => cell.Italic = newValue);
        }

        ImGui.SameLine(0, 4f);
        if (this.DrawToolButton("fmt-underline", buttonSize, "Underline", DrawUnderlineGlyph, styleSourceCell.Underline))
        {
            this.PushUndoSnapshot();
            var newValue = !styleSourceCell.Underline;
            this.ApplyToSelectedCells(cell => cell.Underline = newValue);
        }

        ImGui.SameLine(0, 4f);
        var mergeActive = false;
        var mergeStateKey = primaryKey;
        if (activeTab.TryGetCell(primaryKey, out var mergeStateCell) && mergeStateCell is not null && !string.IsNullOrWhiteSpace(mergeStateCell.MergedInto))
        {
            mergeStateKey = mergeStateCell.MergedInto!;
        }
        if (activeTab.TryGetCell(mergeStateKey, out var mergeRootCell) && mergeRootCell is not null)
        {
            mergeActive = mergeRootCell.RowSpan > 1 || mergeRootCell.ColSpan > 1;
        }
        if (this.DrawToolButton("fmt-merge", buttonSize, "Merge selected cells", DrawMergeGlyph, mergeActive))
        {
            this.MergeSelectedCells();
        }

        ImGui.SameLine(0, 4f);
        if (this.DrawToolButton("fmt-lock", buttonSize, styleSourceCell.Locked ? "Unlock selected cells" : "Lock selected cells", DrawLockGlyph, styleSourceCell.Locked))
        {
            this.PushUndoSnapshot();
            var newLocked = !styleSourceCell.Locked;
            this.ApplyToSelectedCells(cell => cell.Locked = newLocked);
            this.RecordActivity(newLocked ? "Cells locked" : "Cells unlocked", primaryKey, null, null);
        }

        ImGui.SameLine(0, 8f);
        if (this.DrawToolButton("fmt-borders", buttonSize, "Borders", DrawBorderGlyph, styleSourceCell.Borders != CellBorderFlags.None))
        {
            if (this.configuration.LastBorderColor != 0)
            {
                this.selectedBorderColor = ImGui.ColorConvertU32ToFloat4(this.configuration.LastBorderColor);
            }

            ImGui.OpenPopup("fmt-border-popup");
        }
        if (ImGui.BeginPopup("fmt-border-popup"))
        {
            ImGui.TextUnformatted("Border color");
            var borderColor = this.selectedBorderColor;
            if (ImGui.ColorEdit4("##border-color", ref borderColor, ImGuiColorEditFlags.NoInputs | ImGuiColorEditFlags.AlphaBar))
            {
                this.selectedBorderColor = borderColor;
                this.configuration.LastBorderColor = ImGui.ColorConvertFloat4ToU32(this.selectedBorderColor);
                this.configuration.Save();
            }
            var borderColorU32 = ImGui.ColorConvertFloat4ToU32(this.selectedBorderColor);
            if (ImGui.BeginTable("##border-presets", 2, ImGuiTableFlags.SizingStretchSame))
            {
                ImGui.TableNextColumn();
                if (ImGui.Selectable("Top")) { this.PushUndoSnapshot(); this.ApplyBorderPresetToSelection(CellBorderFlags.Top, borderColorU32); }
                if (ImGui.Selectable("Left")) { this.PushUndoSnapshot(); this.ApplyBorderPresetToSelection(CellBorderFlags.Left, borderColorU32); }
                if (ImGui.Selectable("Top/Bottom")) { this.PushUndoSnapshot(); this.ApplyBorderPresetToSelection(CellBorderFlags.Top | CellBorderFlags.Bottom, borderColorU32); }
                if (ImGui.Selectable("All")) { this.PushUndoSnapshot(); this.ApplyBorderPresetToSelection(CellBorderFlags.All, borderColorU32); }
                ImGui.TableNextColumn();
                if (ImGui.Selectable("Right")) { this.PushUndoSnapshot(); this.ApplyBorderPresetToSelection(CellBorderFlags.Right, borderColorU32); }
                if (ImGui.Selectable("Bottom")) { this.PushUndoSnapshot(); this.ApplyBorderPresetToSelection(CellBorderFlags.Bottom, borderColorU32); }
                if (ImGui.Selectable("Sides")) { this.PushUndoSnapshot(); this.ApplyBorderPresetToSelection(CellBorderFlags.Left | CellBorderFlags.Right, borderColorU32); }
                if (ImGui.Selectable("None")) { this.PushUndoSnapshot(); this.ApplyBorderPresetToSelection(CellBorderFlags.None, borderColorU32); }
                ImGui.EndTable();
            }
            ImGui.EndPopup();
        }


        var canUseComments = this.HasSheetPermission(SheetPermissionType.UseComments);
        ImGui.SameLine(0, 4f);
        if (this.DrawToolButton("fmt-comment", buttonSize, canUseComments ? "Add comment" : "You don't have permission to use Add comment", DrawCommentGlyph, false, canUseComments))
        {
            ImGui.OpenPopup("fmt-comment-popup");
        }
        if (ImGui.BeginPopup("fmt-comment-popup"))
        {
            ImGui.TextUnformatted("Add comment:");
            ImGui.SetNextItemWidth(220f);
            ImGui.InputText("##inline-comment", ref this.commentDraft, 512);
            ImGui.SameLine(0, 6f);
            if (this.DrawStyledButton("fmt-comment-add", "Add", AccentGreen, WhiteText, new Vector2(54f, 22f)))
            {
                var feature = this.GetSelectedFeatureBundle(true);
                if (feature is not null && !string.IsNullOrWhiteSpace(this.commentDraft))
                {
                    feature.Comments.Add(new CellComment
                    {
                        AuthorId = this.configuration.UserId,
                        AuthorName = this.GetSafeCharacterFullName(),
                        Message = this.commentDraft.Trim(),
                        CreatedAtUtc = DateTimeOffset.UtcNow,
                    });
                    this.RecordActivity("Comment added", this.GetPrimarySelectedCellKey(), null, this.commentDraft.Trim());
                    this.commentDraft = string.Empty;
                    ImGui.CloseCurrentPopup();
                }
            }
            ImGui.EndPopup();
        }

        ImGui.SameLine(0, 4f);
        if (this.DrawToolButton("fmt-comment-remove", buttonSize, canUseComments ? "Remove your comment" : "You don't have permission to use Remove your comment", DrawCommentDeleteGlyph, false, canUseComments))
        {
            var feature = this.GetSelectedFeatureBundle(false);
            if (feature is not null)
            {
                var removed = feature.Comments.RemoveAll(c => string.Equals(c.AuthorId, this.configuration.UserId, StringComparison.OrdinalIgnoreCase));
                if (removed == 0 && feature.Comments.Count > 0)
                {
                    feature.Comments.RemoveAt(feature.Comments.Count - 1);
                    removed = 1;
                }
                if (removed > 0)
                {
                    this.RecordActivity("Comment removed", this.GetPrimarySelectedCellKey(), null, null);
                }
            }
        }

        if (this.currentRole == SheetAccessRole.Owner && this.currentSheet is not null)
        {
            ImGui.SameLine(0, 12f);
            ImGui.AlignTextToFramePadding();
            ImGui.TextUnformatted("Save Preset");
            ImGui.SameLine(0, 6f);
            ImGui.SetNextItemWidth(170f);
            ImGui.InputText("##SaveSheetPresetName", ref this.saveSheetPresetName, 64);
            ImGui.SameLine(0, 6f);
            var savePresetPressed = this.DrawStyledButton("save-sheet-preset", "✓", HexToVector4("#14891A"), WhiteText, new Vector2(24f, 22f));
            if (savePresetPressed && !string.IsNullOrWhiteSpace(this.saveSheetPresetName))
            {
                this.SaveActiveTabAsPreset(this.saveSheetPresetName);
                this.saveSheetPresetName = string.Empty;
            }
            if (ImGui.IsItemHovered())
            {
                ImGui.SetTooltip("Save the current tab as a reusable preset");
            }
        }

        settings.AutoSaveLocal = false;
        settings.AutoSaveCloud = true;
        settings.AutoSaveSeconds = 20;
    }

    private bool DrawToolButton(string id, Vector2 size, string tooltip, Action<ImDrawListPtr, Vector2, Vector2, bool> drawGlyph, bool active, bool enabled = true)
    {
        if (!enabled)
        {
            ImGui.BeginDisabled();
        }

        var cursor = ImGui.GetCursorScreenPos();
        var clicked = ImGui.InvisibleButton(id, size) && enabled;
        var hovered = ImGui.IsItemHovered();
        if (hovered)
        {
            ImGui.SetMouseCursor(ImGuiMouseCursor.Hand);
        }
        var held = ImGui.IsItemActive();
        if (!enabled)
        {
            ImGui.EndDisabled();
        }

        var baseColor = enabled
            ? (active ? HexToVector4("#4b4b4b") : HexToVector4("#2f2f2f"))
            : HexToVector4("#232323");
        var color = enabled
            ? (held ? AdjustBrightness(baseColor, 0.92f) : hovered ? AdjustBrightness(baseColor, 1.10f) : baseColor)
            : baseColor;
        var min = cursor;
        var max = cursor + size;
        var drawList = ImGui.GetWindowDrawList();
        drawList.AddRectFilled(min, max, ImGui.ColorConvertFloat4ToU32(color), 3f);
        drawList.AddRect(min, max, ImGui.ColorConvertFloat4ToU32(enabled ? (active ? AccentOrange : HexToVector4("#454545")) : HexToVector4("#343434")), 3f);
        drawGlyph(drawList, min, max, enabled && (hovered || active));

        if (hovered)
        {
            ImGui.BeginTooltip();
            ImGui.TextUnformatted(tooltip);
            ImGui.EndTooltip();
        }

        return clicked;
    }

    private bool DrawColorSwatchButton(string id, Vector4 color, Vector2 size)
    {
        var cursor = ImGui.GetCursorScreenPos();
        var clicked = ImGui.InvisibleButton(id, size);
        var hovered = ImGui.IsItemHovered();
        if (hovered)
        {
            ImGui.SetMouseCursor(ImGuiMouseCursor.Hand);
        }
        var drawList = ImGui.GetWindowDrawList();
        var min = cursor;
        var max = cursor + size;
        drawList.AddRectFilled(min + new Vector2(0f, 1f), max + new Vector2(0f, 1f), ImGui.ColorConvertFloat4ToU32(new Vector4(0f, 0f, 0f, 0.16f)), 3f);
        drawList.AddRectFilled(min, max, ImGui.ColorConvertFloat4ToU32(color), 3f);
        drawList.AddRect(min, max, ImGui.ColorConvertFloat4ToU32(hovered ? WhiteText : HexToVector4("#2a2a2a")), 3f, 0, hovered ? 1.2f : 1f);
        if (hovered)
        {
            ImGui.BeginTooltip();
            ImGui.TextUnformatted("Pick sheet color");
            ImGui.EndTooltip();
        }

        return clicked;
    }

    private static void DrawTextColorGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, Vector4 color)
    {
        var w = max.X - min.X;
        var h = max.Y - min.Y;
        var cx = min.X + (w / 2f);
        var top = min.Y + 5f;
        var bottom = max.Y - 7f;
        drawList.AddLine(new Vector2(cx - 4f, bottom), new Vector2(cx, top), ImGui.ColorConvertFloat4ToU32(WhiteText), 1.2f);
        drawList.AddLine(new Vector2(cx + 4f, bottom), new Vector2(cx, top), ImGui.ColorConvertFloat4ToU32(WhiteText), 1.2f);
        drawList.AddLine(new Vector2(cx - 2.2f, min.Y + h * 0.56f), new Vector2(cx + 2.2f, min.Y + h * 0.56f), ImGui.ColorConvertFloat4ToU32(WhiteText), 1.2f);
        drawList.AddLine(new Vector2(min.X + 4f, max.Y - 4f), new Vector2(max.X - 4f, max.Y - 4f), ImGui.ColorConvertFloat4ToU32(color), 2f);
    }

    private static void DrawFillColorGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, Vector4 color)
    {
        var outline = ImGui.ColorConvertFloat4ToU32(WhiteText);
        var fill = ImGui.ColorConvertFloat4ToU32(HexToVector4("#6a6a6a"));
        var indicator = (color.W <= 0.001f && color.X == 0f && color.Y == 0f && color.Z == 0f)
            ? ImGui.ColorConvertFloat4ToU32(WhiteText)
            : ImGui.ColorConvertFloat4ToU32(color);
        var squareMin = new Vector2(min.X + 8f, min.Y + 5f);
        var squareMax = new Vector2(max.X - 8f, min.Y + 12f);
        drawList.AddRectFilled(squareMin, squareMax, fill, 1f);
        drawList.AddRect(squareMin, squareMax, outline, 1f);
        drawList.AddLine(new Vector2(min.X + 5f, max.Y - 5f), new Vector2(max.X - 5f, max.Y - 5f), indicator, 3f);
    }

    private static void DrawAlignGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, int mode)
    {
        var color = ImGui.ColorConvertFloat4ToU32(WhiteText);
        for (var i = 0; i < 3; i++)
        {
            var y = min.Y + 6f + (i * 5f);
            var width = i == 1 ? 9f : 12f;
            var x = mode switch
            {
                1 => min.X + ((max.X - min.X - width) / 2f),
                2 => max.X - width - 5f,
                _ => min.X + 5f,
            };
            drawList.AddLine(new Vector2(x, y), new Vector2(x + width, y), color, 1.2f);
        }
    }


    private static void DrawFontUpGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, bool hovered)
    {
        var color = ImGui.ColorConvertFloat4ToU32(WhiteText);
        drawList.AddText(new Vector2(min.X + 5f, min.Y + 4f), color, "A");
        drawList.AddLine(new Vector2(max.X - 9f, min.Y + 7f), new Vector2(max.X - 5f, min.Y + 7f), color, 1.1f);
        drawList.AddLine(new Vector2(max.X - 7f, min.Y + 5f), new Vector2(max.X - 7f, min.Y + 9f), color, 1.1f);
    }

    private static void DrawFontDownGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, bool hovered)
    {
        var color = ImGui.ColorConvertFloat4ToU32(WhiteText);
        drawList.AddText(new Vector2(min.X + 5f, min.Y + 4f), color, "A");
        drawList.AddLine(new Vector2(max.X - 9f, min.Y + 7f), new Vector2(max.X - 5f, min.Y + 7f), color, 1.1f);
    }

    private static void DrawBoldGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, bool hovered)
    {
        drawList.AddText(new Vector2(min.X + 8f, min.Y + 4f), ImGui.ColorConvertFloat4ToU32(WhiteText), "N");
    }

    private static void DrawItalicGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, bool hovered)
    {
        var color = ImGui.ColorConvertFloat4ToU32(WhiteText);
        var text = "I";
        var textSize = ImGui.CalcTextSize(text);
        var pos = new Vector2(min.X + ((max.X - min.X) - textSize.X) / 2f + 1f, min.Y + ((max.Y - min.Y) - textSize.Y) / 2f - 1f);
        drawList.AddText(pos, color, text);
        drawList.AddText(pos + new Vector2(0.4f, 0f), color, text);
    }

    private static void DrawUnderlineGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, bool hovered)
    {
        var color = ImGui.ColorConvertFloat4ToU32(WhiteText);
        var text = "S";
        var textSize = ImGui.CalcTextSize(text);
        var pos = new Vector2(min.X + ((max.X - min.X) - textSize.X) / 2f, min.Y + 3f);
        drawList.AddText(pos, color, text);
        drawList.AddLine(new Vector2(min.X + 7f, max.Y - 5f), new Vector2(max.X - 7f, max.Y - 5f), color, 1.2f);
    }

    private static void DrawMergeGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, bool hovered)
    {
        var color = ImGui.ColorConvertFloat4ToU32(WhiteText);
        var text = "⇔";
        var textSize = ImGui.CalcTextSize(text);
        var pos = new Vector2(min.X + ((max.X - min.X) - textSize.X) / 2f, min.Y + ((max.Y - min.Y) - textSize.Y) / 2f - 1f);
        drawList.AddText(pos, color, text);
        drawList.AddText(pos + new Vector2(0.6f, 0f), color, text);
    }

    private static void DrawBorderGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, bool hovered)
    {
        var color = ImGui.ColorConvertFloat4ToU32(WhiteText);
        var innerMin = new Vector2(min.X + 6f, min.Y + 5f);
        var innerMax = new Vector2(max.X - 6f, max.Y - 5f);
        drawList.AddRect(innerMin, innerMax, color, 0f, 0, 1.1f);
        drawList.AddLine(new Vector2(innerMin.X, (innerMin.Y + innerMax.Y) / 2f), new Vector2(innerMax.X, (innerMin.Y + innerMax.Y) / 2f), color, 1.1f);
        drawList.AddLine(new Vector2((innerMin.X + innerMax.X) / 2f, innerMin.Y), new Vector2((innerMin.X + innerMax.X) / 2f, innerMax.Y), color, 1.1f);
    }

    private static void DrawLockGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, bool hovered)
    {
        var color = ImGui.ColorConvertFloat4ToU32(WhiteText);
        var cx = (min.X + max.X) / 2f;
        drawList.AddRect(new Vector2(cx - 5f, min.Y + 11f), new Vector2(cx + 5f, max.Y - 5f), color, 1f, 0, 1.1f);
        drawList.AddBezierCubic(new Vector2(cx - 4f, min.Y + 12f), new Vector2(cx - 4f, min.Y + 5f), new Vector2(cx + 4f, min.Y + 5f), new Vector2(cx + 4f, min.Y + 12f), color, 1.1f);
    }

    private static void DrawCommentGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, bool hovered)
    {
        var color = ImGui.ColorConvertFloat4ToU32(WhiteText);
        var bubbleMin = new Vector2(min.X + 5f, min.Y + 5f);
        var bubbleMax = new Vector2(max.X - 5f, max.Y - 7f);
        drawList.AddRect(bubbleMin, bubbleMax, color, 2f, 0, 1.1f);
        drawList.AddTriangle(new Vector2(bubbleMin.X + 5f, bubbleMax.Y), new Vector2(bubbleMin.X + 9f, bubbleMax.Y), new Vector2(bubbleMin.X + 6f, bubbleMax.Y + 4f), color, 1.1f);
    }

    private static void DrawCommentDeleteGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, bool hovered)
    {
        DrawCommentGlyph(drawList, min, max, hovered);
        var color = ImGui.ColorConvertFloat4ToU32(WhiteText);
        var x1 = max.X - 10f;
        var y1 = min.Y + 6f;
        var x2 = max.X - 6f;
        var y2 = min.Y + 10f;
        drawList.AddLine(new Vector2(x1, y1), new Vector2(x2, y2), color, 1.2f);
        drawList.AddLine(new Vector2(x2, y1), new Vector2(x1, y2), color, 1.2f);
    }

    private static void DrawSaveLetterGlyph(ImDrawListPtr drawList, Vector2 min, Vector2 max, string letter)
    {
        var color = ImGui.ColorConvertFloat4ToU32(WhiteText);
        drawList.AddRect(new Vector2(min.X + 6f, min.Y + 5f), new Vector2(max.X - 6f, max.Y - 5f), color, 1f, 0, 1.1f);
        drawList.AddLine(new Vector2(min.X + 8f, min.Y + 8f), new Vector2(max.X - 8f, min.Y + 8f), color, 1.1f);
        drawList.AddText(new Vector2(min.X + 10f, min.Y + 7f), color, letter);
    }

    private readonly record struct FormulaResult(bool Success, double Value, string Error)
    {
        public static FormulaResult Ok(double value) => new(true, value, string.Empty);

        public static FormulaResult Fail(string error) => new(false, 0d, error);
    }

    private sealed class FormulaParser
    {
        private readonly string text;
        private readonly Func<string, FormulaResult> resolveReference;
        private int index;

        public FormulaParser(string text, Func<string, FormulaResult> resolveReference)
        {
            this.text = text;
            this.resolveReference = resolveReference;
        }

        public FormulaResult Parse()
        {
            var result = this.ParseExpression();
            this.SkipWhitespace();
            return !result.Success
                ? result
                : this.index < this.text.Length
                    ? FormulaResult.Fail("#ERR")
                    : result;
        }

        private FormulaResult ParseExpression()
        {
            var left = this.ParseTerm();
            if (!left.Success)
            {
                return left;
            }

            while (true)
            {
                this.SkipWhitespace();
                if (!this.Match('+') && !this.Match('-'))
                {
                    return left;
                }

                var op = this.text[this.index - 1];
                var right = this.ParseTerm();
                if (!right.Success)
                {
                    return right;
                }

                left = op == '+' ? FormulaResult.Ok(left.Value + right.Value) : FormulaResult.Ok(left.Value - right.Value);
            }
        }

        private FormulaResult ParseTerm()
        {
            var left = this.ParseFactor();
            if (!left.Success)
            {
                return left;
            }

            while (true)
            {
                this.SkipWhitespace();
                if (!this.Match('*') && !this.Match('/'))
                {
                    return left;
                }

                var op = this.text[this.index - 1];
                var right = this.ParseFactor();
                if (!right.Success)
                {
                    return right;
                }

                if (op == '/' && Math.Abs(right.Value) < 0.0000001d)
                {
                    return FormulaResult.Fail("#DIV/0");
                }

                left = op == '*' ? FormulaResult.Ok(left.Value * right.Value) : FormulaResult.Ok(left.Value / right.Value);
            }
        }

        private FormulaResult ParseFactor()
        {
            this.SkipWhitespace();
            if (this.Match('+'))
            {
                return this.ParseFactor();
            }

            if (this.Match('-'))
            {
                var nested = this.ParseFactor();
                return nested.Success ? FormulaResult.Ok(-nested.Value) : nested;
            }

            if (this.Match('('))
            {
                var inner = this.ParseExpression();
                if (!inner.Success)
                {
                    return inner;
                }

                this.SkipWhitespace();
                return this.Match(')') ? inner : FormulaResult.Fail("#ERR");
            }

            if (this.index < this.text.Length && char.IsLetter(this.text[this.index]))
            {
                var start = this.index;
                while (this.index < this.text.Length && char.IsLetter(this.text[this.index]))
                {
                    this.index++;
                }

                while (this.index < this.text.Length && char.IsDigit(this.text[this.index]))
                {
                    this.index++;
                }

                var reference = this.text[start..this.index];
                return this.resolveReference(reference);
            }

            return this.ParseNumber();
        }

        private FormulaResult ParseNumber()
        {
            this.SkipWhitespace();
            var start = this.index;
            var seenDecimal = false;
            while (this.index < this.text.Length)
            {
                var ch = this.text[this.index];
                if (char.IsDigit(ch))
                {
                    this.index++;
                    continue;
                }

                if (ch == '.' && !seenDecimal)
                {
                    seenDecimal = true;
                    this.index++;
                    continue;
                }

                break;
            }

            if (start == this.index)
            {
                return FormulaResult.Fail("#ERR");
            }

            var slice = this.text[start..this.index];
            return double.TryParse(slice, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
                ? FormulaResult.Ok(value)
                : FormulaResult.Fail("#ERR");
        }

        private void SkipWhitespace()
        {
            while (this.index < this.text.Length && char.IsWhiteSpace(this.text[this.index]))
            {
                this.index++;
            }
        }

        private bool Match(char target)
        {
            if (this.index < this.text.Length && this.text[this.index] == target)
            {
                this.index++;
                return true;
            }

            return false;
        }
    }
}
