using System.Diagnostics;
using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using EZSheets.Models;

namespace EZSheets.Services;

public sealed class SupabaseRestClient : IDisposable
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = false,
    };

    private readonly Configuration configuration;
    private readonly HttpClient httpClient;
    private readonly SemaphoreSlim discordLoginGate = new(1, 1);
    private readonly SemaphoreSlim refreshTokenGate = new(1, 1);
    private readonly SemaphoreSlim cloudRequestGate = new(1, 1);
    private DateTimeOffset suppressTokenRefreshUntilUtc = DateTimeOffset.MinValue;
    private readonly CancellationTokenSource disposeCts = new();
    private CancellationTokenSource? activeDiscordLoginCts;
    private TcpListener? activeDiscordLoginListener;
    private bool activeDiscordLoginCancelledByUser;
    private string? loadedCharacterSessionKey;
    private string? loadedCharacterFullName;

    public SupabaseRestClient(Configuration configuration)
    {
        this.configuration = configuration;
        this.httpClient = new HttpClient
        {
            Timeout = Timeout.InfiniteTimeSpan,
        };
        this.ApplyBaseAddress();
    }

    public bool HasConfiguredProject
        => !string.IsNullOrWhiteSpace(this.EffectiveSupabaseUrl)
           && !string.IsNullOrWhiteSpace(this.EffectivePublishableKey)
           && !string.IsNullOrWhiteSpace(this.configuration.DiscordClientId);

    public bool IsAuthenticated
    {
        get
        {
            this.EnsureCharacterSessionContextLoaded();
            return (!string.IsNullOrWhiteSpace(this.configuration.ServerSessionToken)
                    || !string.IsNullOrWhiteSpace(this.configuration.AccessToken))
                   && !string.IsNullOrWhiteSpace(this.configuration.UserId);
        }
    }

    public string EffectiveSupabaseUrl => Configuration.NormalizeSupabaseUrl(this.configuration.SupabaseUrl);

    public string EffectivePublishableKey
        => string.IsNullOrWhiteSpace(this.configuration.SupabasePublicApiKey)
            ? Configuration.EmbeddedSupabasePublishableKey
            : this.configuration.SupabasePublicApiKey.Trim();

    public bool IsDiscordLoginInProgress => this.activeDiscordLoginCts is not null;

    public void CancelActiveDiscordLogin()
    {
        this.activeDiscordLoginCancelledByUser = true;

        try
        {
            this.activeDiscordLoginCts?.Cancel();
        }
        catch
        {
            // Ignore cancellation races.
        }

        try
        {
            this.activeDiscordLoginListener?.Stop();
        }
        catch
        {
            // Ignore cleanup races.
        }
    }

    public void ReloadConfiguration()
    {
        this.ApplyBaseAddress();
        this.loadedCharacterSessionKey = null;
        this.loadedCharacterFullName = null;
    }

    public async Task<DiscordAppSession> SignInWithDiscordAsync(CancellationToken cancellationToken = default)
    {
        this.EnsureCharacterSessionContextLoaded();

        var rememberedRefresh = this.GetCurrentRefreshToken();
        if (!string.IsNullOrWhiteSpace(rememberedRefresh))
        {
            try
            {
                var restored = await this.RefreshDiscordSessionAsync(cancellationToken).ConfigureAwait(false);
                this.StoreSession(restored);
                return restored;
            }
            catch
            {
                this.configuration.ClearCharacterSession(this.loadedCharacterFullName, keepRememberedLogin: false);
                this.configuration.Save();
            }
        }

        await this.discordLoginGate.WaitAsync(cancellationToken).ConfigureAwait(false);

        CancellationTokenSource? loginCts = null;
        TcpListener? listener = null;

        try
        {
            var redirectUrl = Configuration.DiscordCallbackUrl;
            var oauthState = GenerateStateToken();

            loginCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, this.disposeCts.Token);
            loginCts.CancelAfter(TimeSpan.FromMinutes(5));
            this.activeDiscordLoginCancelledByUser = false;
            this.activeDiscordLoginCts = loginCts;

            listener = new TcpListener(IPAddress.Loopback, Configuration.DiscordCallbackPort);
            this.activeDiscordLoginListener = listener;

            try
            {
                listener.Start(1);
            }
            catch (SocketException ex)
            {
                throw new InvalidOperationException($"EZSheets could not open the local login callback on port {Configuration.DiscordCallbackPort}. Close any previous login attempt, or restart the game if another plugin instance is still holding that port.", ex);
            }

            var authUrl = this.BuildDiscordAuthorizeUrl(redirectUrl, oauthState);
            OpenInBrowser(authUrl);

            while (true)
            {
                using var client = await listener.AcceptTcpClientAsync(loginCts.Token).ConfigureAwait(false);
                var callback = await ReadLoopbackRequestAsync(client, loginCts.Token).ConfigureAwait(false);

                if (!callback.Path.Equals("/callback/", StringComparison.OrdinalIgnoreCase)
                    && !callback.Path.Equals("/callback", StringComparison.OrdinalIgnoreCase))
                {
                    await WriteHtmlAsync(client.GetStream(), "Not found", "EZSheets received a request for an unexpected path.", HttpStatusCode.NotFound).ConfigureAwait(false);
                    continue;
                }

                var parameters = callback.Parameters;

                if (parameters.TryGetValue("error_description", out var errorDescription) && !string.IsNullOrWhiteSpace(errorDescription))
                {
                    await WriteHtmlAsync(client.GetStream(), "Discord sign-in failed", "EZSheets received an OAuth error. You can close this window and return to the game.").ConfigureAwait(false);
                    throw new InvalidOperationException(Uri.UnescapeDataString(errorDescription));
                }

                if (parameters.TryGetValue("error", out var errorCode) && !string.IsNullOrWhiteSpace(errorCode))
                {
                    await WriteHtmlAsync(client.GetStream(), "Discord sign-in failed", "EZSheets received an OAuth error. You can close this window and return to the game.").ConfigureAwait(false);
                    throw new InvalidOperationException(Uri.UnescapeDataString(errorCode));
                }

                if (!parameters.TryGetValue("state", out var returnedState) || string.IsNullOrWhiteSpace(returnedState))
                {
                    await WriteHtmlAsync(client.GetStream(), "Missing state", "Discord returned without a valid state value. You can close this window.").ConfigureAwait(false);
                    throw new InvalidOperationException("Discord login did not return a valid state value.");
                }

                if (!string.Equals(returnedState, oauthState, StringComparison.Ordinal))
                {
                    await WriteHtmlAsync(client.GetStream(), "Invalid state", "Discord returned an invalid state value. You can close this window.").ConfigureAwait(false);
                    throw new InvalidOperationException("Discord login returned an invalid state value.");
                }

                if (!parameters.TryGetValue("code", out var code) || string.IsNullOrWhiteSpace(code))
                {
                    await WriteHtmlAsync(client.GetStream(), "Missing code", "The browser returned without an authorization code. You can close this window.").ConfigureAwait(false);
                    throw new InvalidOperationException("Discord login did not return an authorization code.");
                }

                var session = await this.ExchangeDiscordCodeForSessionAsync(code, redirectUrl, loginCts.Token).ConfigureAwait(false);
                this.StoreSession(session);
                await WriteHtmlAsync(client.GetStream(), "EZSheets login complete", "Discord login succeeded. You can close this browser tab and return to the game.").ConfigureAwait(false);
                return session;
            }
        }
        catch (OperationCanceledException) when (this.disposeCts.IsCancellationRequested)
        {
            throw new InvalidOperationException("EZSheets cancelled the Discord login because the plugin was reloaded or closed.");
        }
        catch (OperationCanceledException)
        {
            if (this.activeDiscordLoginCancelledByUser)
            {
                throw new InvalidOperationException("Discord login was cancelled.");
            }

            throw new InvalidOperationException("Discord login timed out before the browser returned to EZSheets.");
        }
        finally
        {
            try
            {
                listener?.Stop();
            }
            catch
            {
                // Ignore cleanup failures.
            }

            if (ReferenceEquals(this.activeDiscordLoginListener, listener))
            {
                this.activeDiscordLoginListener = null;
            }

            if (ReferenceEquals(this.activeDiscordLoginCts, loginCts))
            {
                this.activeDiscordLoginCts = null;
            }

            this.activeDiscordLoginCancelledByUser = false;
            loginCts?.Dispose();
            this.discordLoginGate.Release();
        }
    }

    public async Task<bool> RestoreSessionAsync(CancellationToken cancellationToken = default)
    {
        this.EnsureCharacterSessionContextLoaded();

        var refreshToken = this.GetCurrentRefreshToken();
        if (!this.HasConfiguredProject || string.IsNullOrWhiteSpace(refreshToken))
        {
            return false;
        }

        try
        {
            try
            {
                var session = await this.RefreshDiscordSessionAsync(cancellationToken).ConfigureAwait(false);
                this.StoreSession(session);
            }
            catch (SupabaseApiException ex) when (string.Equals(ex.Message, "invalid_grant", StringComparison.OrdinalIgnoreCase))
            {
                this.configuration.ClearCharacterSession(this.loadedCharacterFullName, keepRememberedLogin: false);
                this.configuration.Save();
                throw new InvalidOperationException("Your Discord login expired. Sign in again to continue using cloud features.", ex);
            }
            return true;
        }
        catch
        {
            this.configuration.ClearCharacterSession(this.loadedCharacterFullName, keepRememberedLogin: false);
            this.configuration.Save();
            return false;
        }
    }

    public async Task SignOutAsync(CancellationToken cancellationToken = default)
    {
        this.EnsureCharacterSessionContextLoaded();

        _ = cancellationToken;
        this.configuration.ClearCharacterSession(this.loadedCharacterFullName, keepRememberedLogin: true);
        this.configuration.Save();
        await Task.CompletedTask.ConfigureAwait(false);
    }

    public async Task<List<SheetSummary>> ListSheetsAsync(CancellationToken cancellationToken = default)
    {
        var rows = await this.SendAuthenticatedFunctionAsync<List<SheetSummary>>("list_sheets", null, cancellationToken).ConfigureAwait(false);
        return rows ?? new List<SheetSummary>();
    }

    public async Task<RemoteSheet?> GetSheetAsync(string sheetId, CancellationToken cancellationToken = default)
    {
        try
        {
            var remote = await this.SendAuthenticatedFunctionAsync<RemoteSheet>(
                "get_sheet",
                new { sheet_id = sheetId },
                cancellationToken).ConfigureAwait(false);

            remote.Data ??= new SheetDocument();
            remote.Data.Normalize(remote.RowsCount, remote.ColsCount);
            return remote;
        }
        catch (SupabaseApiException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
        {
            return null;
        }
    }

    public async Task<SheetAccessRole> GetAccessRoleAsync(string sheetId, CancellationToken cancellationToken = default)
    {
        var response = await this.SendAuthenticatedFunctionAsync<AccessRoleResponse>(
            "get_access_role",
            new { sheet_id = sheetId },
            cancellationToken).ConfigureAwait(false);

        var role = response.Role?.Trim().ToLowerInvariant();
        return role switch
        {
            "owner" => SheetAccessRole.Owner,
            "editor" => SheetAccessRole.Editor,
            _ => SheetAccessRole.Viewer,
        };
    }

    public async Task<RemoteSheet> CreateSheetAsync(string title, int rowsCount, int colsCount, string defaultRole, SheetDocument data, CancellationToken cancellationToken = default)
    {
        this.EnsureAuthenticated();
        data.Normalize(rowsCount, colsCount);

        for (var attempt = 0; attempt < 5; attempt++)
        {
            var code = GenerateShareCode();
            try
            {
                var created = await this.SendAuthenticatedFunctionAsync<RemoteSheet>(
                    "create_sheet",
                    new
                    {
                        title,
                        code,
                        rows_count = rowsCount,
                        cols_count = colsCount,
                        default_role = defaultRole,
                        data,
                        version = 1L,
                        character_name = this.TryGetCurrentCharacterFullName(),
                    },
                    cancellationToken).ConfigureAwait(false);

                created.Data.Normalize(created.RowsCount, created.ColsCount);
                return created;
            }
            catch (SupabaseApiException ex) when (ex.StatusCode == HttpStatusCode.Conflict || ex.Message.Contains("duplicate", StringComparison.OrdinalIgnoreCase))
            {
                // Share code collision; try again.
            }
        }

        throw new InvalidOperationException("Could not generate a unique share code after multiple attempts.");
    }

    public async Task<RemoteSheet> UpdateSheetAsync(RemoteSheet sheet, CancellationToken cancellationToken = default)
    {
        this.EnsureAuthenticated();
        sheet.Data.Normalize(sheet.RowsCount, sheet.ColsCount);

        try
        {
            var payloadData = BuildPersistableSheetPayload(sheet, aggressiveTrim: false);
            return await this.SendAuthenticatedFunctionAsync<RemoteSheet>(
                "update_sheet",
                new
                {
                    id = sheet.Id,
                    title = sheet.Title,
                    rows_count = sheet.RowsCount,
                    cols_count = sheet.ColsCount,
                    default_role = sheet.DefaultRole,
                    data = payloadData,
                    version = sheet.Version,
                },
                cancellationToken).ConfigureAwait(false);
        }
        catch (SupabaseApiException ex) when (ex.Message.Contains("statement timeout", StringComparison.OrdinalIgnoreCase))
        {
            var payloadData = BuildPersistableSheetPayload(sheet, aggressiveTrim: true);
            return await this.SendAuthenticatedFunctionAsync<RemoteSheet>(
                "update_sheet",
                new
                {
                    id = sheet.Id,
                    title = sheet.Title,
                    rows_count = sheet.RowsCount,
                    cols_count = sheet.ColsCount,
                    default_role = sheet.DefaultRole,
                    data = payloadData,
                    version = sheet.Version,
                },
                cancellationToken).ConfigureAwait(false);
        }
    }

    private static SheetDocument BuildPersistableSheetPayload(RemoteSheet sheet, bool aggressiveTrim)
    {
        return SheetSerializationHelper.CloneForPersistence(sheet.Data, sheet.RowsCount, sheet.ColsCount, aggressiveTrim);
    }

    private static List<T> TrimTail<T>(List<T>? source, int maxCount)
    {
        if (source is null || source.Count == 0)
        {
            return new List<T>();
        }

        if (source.Count <= maxCount)
        {
            return new List<T>(source);
        }

        return source.Skip(Math.Max(0, source.Count - maxCount)).ToList();
    }

    public async Task<string> JoinSheetByCodeAsync(string code, CancellationToken cancellationToken = default)
    {
        this.EnsureAuthenticated();

        var response = await this.SendAuthenticatedFunctionAsync<JoinSheetResponse>(
            "join_sheet_by_code",
            new
            {
                code = code.Trim().ToUpperInvariant(),
                character_name = this.TryGetCurrentCharacterFullName(),
            },
            cancellationToken).ConfigureAwait(false);

        if (string.IsNullOrWhiteSpace(response.SheetId))
        {
            throw new InvalidOperationException("EZSheets did not return a sheet id for the join operation.");
        }

        return response.SheetId;
    }

    public async Task DeleteSheetAsync(string sheetId, CancellationToken cancellationToken = default)
    {
        this.EnsureAuthenticated();
        await this.SendAuthenticatedFunctionAsync<FunctionSuccessResponse>(
            "delete_sheet",
            new { sheet_id = sheetId },
            cancellationToken).ConfigureAwait(false);
    }


    public async Task<List<SheetMemberRow>> ListSheetMembersAsync(string sheetId, CancellationToken cancellationToken = default)
    {
        this.EnsureAuthenticated();

        var rows = await this.SendAuthenticatedFunctionAsync<List<SheetMemberRow>>(
            "list_sheet_members",
            new { sheet_id = sheetId },
            cancellationToken).ConfigureAwait(false);

        return rows ?? new List<SheetMemberRow>();
    }

    public async Task RemoveSheetMemberAsync(string sheetId, string userId, string? reason, CancellationToken cancellationToken = default)
    {
        this.EnsureAuthenticated();

        await this.SendAuthenticatedFunctionAsync<FunctionSuccessResponse>(
            "remove_sheet_member",
            new
            {
                sheet_id = sheetId,
                user_id = userId,
                reason,
            },
            cancellationToken).ConfigureAwait(false);
    }

    public async Task SyncSheetPresenceAsync(string sheetId, string characterName, string? activeTabName, string? editingCellKey, CancellationToken cancellationToken = default)
    {
        this.EnsureAuthenticated();
        await this.SendAuthenticatedFunctionAsync<FunctionSuccessResponse>(
            "sync_presence",
            new
            {
                sheet_id = sheetId,
                character_name = characterName,
                active_tab_name = activeTabName,
                editing_cell_key = editingCellKey,
            },
            cancellationToken).ConfigureAwait(false);
    }

    public async Task PostChatMessageAsync(string sheetId, string characterName, string message, CancellationToken cancellationToken = default)
    {
        this.EnsureAuthenticated();
        await this.SendAuthenticatedFunctionAsync<FunctionSuccessResponse>(
            "post_chat_message",
            new
            {
                sheet_id = sheetId,
                character_name = characterName,
                message,
            },
            cancellationToken).ConfigureAwait(false);
    }


    public async Task<SheetRuntimeState> GetSheetRuntimeStateAsync(string sheetId, CancellationToken cancellationToken = default)
    {
        this.EnsureAuthenticated();
        return await this.SendAuthenticatedFunctionAsync<SheetRuntimeState>(
            "get_sheet_runtime",
            new { sheet_id = sheetId, include_members = false },
            cancellationToken).ConfigureAwait(false);
    }

    public async Task AcquireCellLockAsync(string sheetId, string cellKey, string characterName, CancellationToken cancellationToken = default)
    {
        this.EnsureAuthenticated();
        await this.SendAuthenticatedFunctionAsync<FunctionSuccessResponse>(
            "acquire_cell_lock",
            new
            {
                sheet_id = sheetId,
                cell_key = cellKey,
                character_name = characterName,
            },
            cancellationToken).ConfigureAwait(false);
    }

    public async Task ReleaseCellLockAsync(string sheetId, string? cellKey, CancellationToken cancellationToken = default)
    {
        this.EnsureAuthenticated();
        await this.SendAuthenticatedFunctionAsync<FunctionSuccessResponse>(
            "release_cell_lock",
            new
            {
                sheet_id = sheetId,
                cell_key = cellKey,
            },
            cancellationToken).ConfigureAwait(false);
    }

    public async Task<SheetUniqueCodeInfo> CreateUniqueCodeAsync(string sheetId, bool invalidateCurrent, CancellationToken cancellationToken = default)
    {
        this.EnsureAuthenticated();
        return await this.SendAuthenticatedFunctionAsync<SheetUniqueCodeInfo>(
            invalidateCurrent ? "invalidate_and_create_unique_code" : "create_unique_code",
            new { sheet_id = sheetId },
            cancellationToken).ConfigureAwait(false);
    }

    public async Task<List<SheetBlockedMemberRow>> ListSheetBlocklistAsync(string sheetId, CancellationToken cancellationToken = default)
    {
        this.EnsureAuthenticated();
        var rows = await this.SendAuthenticatedFunctionAsync<List<SheetBlockedMemberRow>>(
            "list_sheet_blocklist",
            new { sheet_id = sheetId },
            cancellationToken).ConfigureAwait(false);
        return rows ?? new List<SheetBlockedMemberRow>();
    }

    public async Task UnblockSheetMemberAsync(string sheetId, string userId, CancellationToken cancellationToken = default)
    {
        this.EnsureAuthenticated();
        await this.SendAuthenticatedFunctionAsync<FunctionSuccessResponse>(
            "unblock_sheet_member",
            new
            {
                sheet_id = sheetId,
                user_id = userId,
            },
            cancellationToken).ConfigureAwait(false);
    }

    private void EnsureCharacterSessionContextLoaded()
    {
        var characterFullName = this.TryGetCurrentCharacterFullName();
        var sessionKey = Configuration.NormalizeCharacterSessionKey(characterFullName);
        if (string.IsNullOrWhiteSpace(sessionKey))
        {
            if (!string.IsNullOrWhiteSpace(this.loadedCharacterSessionKey))
            {
                this.loadedCharacterSessionKey = null;
                this.loadedCharacterFullName = null;
                this.configuration.ClearSession();
            }

            return;
        }

        if (string.Equals(this.loadedCharacterSessionKey, sessionKey, StringComparison.OrdinalIgnoreCase))
        {
            this.loadedCharacterFullName = characterFullName;
            if (!string.IsNullOrWhiteSpace(characterFullName) && !string.Equals(this.configuration.LastCharacterFullName, characterFullName, StringComparison.Ordinal))
            {
                this.configuration.LastCharacterFullName = characterFullName;
            }

            return;
        }

        this.loadedCharacterSessionKey = sessionKey;
        this.loadedCharacterFullName = characterFullName;
        this.configuration.TryLoadCharacterSession(characterFullName);
    }

    private string TryGetCurrentCharacterFullName()
    {
        try
        {
            var localPlayer = Plugin.ObjectTable?.LocalPlayer;
            var rawName = localPlayer?.Name.TextValue?.Trim() ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(rawName))
            {
                return rawName;
            }
        }
        catch
        {
            // Ignore lookup failures and fall back to the cached character name.
        }

        return this.configuration.LastCharacterFullName?.Trim() ?? string.Empty;
    }

    private string GetCurrentRefreshToken()
    {
        return !string.IsNullOrWhiteSpace(this.configuration.RefreshToken)
            ? this.configuration.RefreshToken
            : this.configuration.RememberedRefreshToken;
    }

    private bool HasValidServerSession()
    {
        this.EnsureCharacterSessionContextLoaded();
        if (string.IsNullOrWhiteSpace(this.configuration.ServerSessionToken))
        {
            return false;
        }

        if (this.configuration.ServerSessionExpiresAtUnix <= 0)
        {
            return true;
        }

        return DateTimeOffset.UtcNow.ToUnixTimeSeconds() < this.configuration.ServerSessionExpiresAtUnix - 20;
    }

    private void PersistCurrentCharacterSession()
    {
        this.EnsureCharacterSessionContextLoaded();
        this.configuration.SaveCurrentSessionToCharacter(this.loadedCharacterFullName);
        this.configuration.Save();
    }

    public void Dispose()
    {
        try
        {
            this.disposeCts.Cancel();
        }
        catch
        {
            // Ignore cancellation failures during shutdown.
        }

        try
        {
            this.activeDiscordLoginListener?.Stop();
        }
        catch
        {
            // Ignore cleanup failures during shutdown.
        }

        this.activeDiscordLoginCts?.Dispose();
        this.disposeCts.Dispose();
        this.discordLoginGate.Dispose();
        this.refreshTokenGate.Dispose();
        this.cloudRequestGate.Dispose();
        this.httpClient.Dispose();
    }

    private void ApplyBaseAddress()
    {
        if (Uri.TryCreate(this.EffectiveSupabaseUrl.TrimEnd('/') + "/", UriKind.Absolute, out var baseAddress))
        {
            this.httpClient.BaseAddress = baseAddress;
        }
    }

    private async Task EnsureFreshDiscordAccessTokenAsync(CancellationToken cancellationToken)
    {
        if (!this.IsAuthenticated)
        {
            return;
        }

        if (this.HasValidServerSession())
        {
            return;
        }

        if (this.configuration.AccessTokenExpiresAtUnix <= 0)
        {
            return;
        }

        var nowUnix = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
        if (nowUnix < this.configuration.AccessTokenExpiresAtUnix - 20)
        {
            return;
        }

        if (this.suppressTokenRefreshUntilUtc > DateTimeOffset.UtcNow)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(this.GetCurrentRefreshToken()))
        {
            throw new InvalidOperationException("Your Discord session expired. Please sign in again.");
        }

        await this.refreshTokenGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            nowUnix = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
            if (this.configuration.AccessTokenExpiresAtUnix > 0 && nowUnix < this.configuration.AccessTokenExpiresAtUnix - 20)
            {
                return;
            }

            try
            {
                var session = await this.RefreshDiscordSessionAsync(cancellationToken, alreadySerialized: true).ConfigureAwait(false);
                this.StoreSession(session);
                this.suppressTokenRefreshUntilUtc = DateTimeOffset.MinValue;
            }
            catch (SupabaseApiException ex) when (string.Equals(ex.Message, "invalid_grant", StringComparison.OrdinalIgnoreCase))
            {
                if (this.configuration.TryReloadCharacterSessionFromPersisted(this.loadedCharacterFullName))
                {
                    this.suppressTokenRefreshUntilUtc = DateTimeOffset.MinValue;
                    return;
                }

                // Keep using the current access token until the server actually rejects it.
                // This avoids freezing cloud features early just because the refresh token rotated or expired.
                this.suppressTokenRefreshUntilUtc = DateTimeOffset.UtcNow.AddMinutes(2);
                return;
            }
        }
        finally
        {
            this.refreshTokenGate.Release();
        }
    }

    private static bool RequiresCloudSerialization(string action)
        => action switch
        {
            "create_sheet" => true,
            "update_sheet" => true,
            "delete_sheet" => true,
            "remove_sheet_member" => true,
            "unblock_sheet_member" => true,
            "create_unique_code" => true,
            "invalidate_and_create_unique_code" => true,
            _ => false,
        };

    private async Task<T> SendAuthenticatedFunctionAsync<T>(string action, object? payload, CancellationToken cancellationToken, bool allowRefreshRetry = true, bool allowPersistedRetry = true, bool alreadySerialized = false, int retryAttempt = 0)
        where T : class
    {
        if (!alreadySerialized)
        {
            if (RequiresCloudSerialization(action))
            {
                await this.cloudRequestGate.WaitAsync(cancellationToken).ConfigureAwait(false);
                try
                {
                    return await this.SendAuthenticatedFunctionAsync<T>(action, payload, cancellationToken, allowRefreshRetry, allowPersistedRetry, alreadySerialized: true, retryAttempt).ConfigureAwait(false);
                }
                finally
                {
                    this.cloudRequestGate.Release();
                }
            }

            return await this.SendAuthenticatedFunctionAsync<T>(action, payload, cancellationToken, allowRefreshRetry, allowPersistedRetry, alreadySerialized: true, retryAttempt).ConfigureAwait(false);
        }

        await this.EnsureFreshDiscordAccessTokenAsync(cancellationToken).ConfigureAwait(false);
        this.EnsureAuthenticated();

        using var request = this.CreateFunctionRequest(action, payload, includeAuth: true);
        using var requestCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, this.disposeCts.Token);
        requestCts.CancelAfter(GetActionTimeout(action));

        HttpResponseMessage response;
        try
        {
            response = await this.httpClient.SendAsync(request, requestCts.Token).ConfigureAwait(false);
        }
        catch (TaskCanceledException ex) when (!cancellationToken.IsCancellationRequested && !this.disposeCts.IsCancellationRequested)
        {
            if (ShouldRetryAction(action, retryAttempt))
            {
                await Task.Delay(GetRetryDelay(retryAttempt), cancellationToken).ConfigureAwait(false);
                return await this.SendAuthenticatedFunctionAsync<T>(action, payload, cancellationToken, allowRefreshRetry, allowPersistedRetry, alreadySerialized: true, retryAttempt + 1).ConfigureAwait(false);
            }

            throw new TimeoutException("EZSheets cloud request timed out. Try again in a moment.", ex);
        }
        catch (HttpRequestException ex) when (!cancellationToken.IsCancellationRequested && !this.disposeCts.IsCancellationRequested)
        {
            if (ShouldRetryAction(action, retryAttempt))
            {
                await Task.Delay(GetRetryDelay(retryAttempt), cancellationToken).ConfigureAwait(false);
                return await this.SendAuthenticatedFunctionAsync<T>(action, payload, cancellationToken, allowRefreshRetry, allowPersistedRetry, alreadySerialized: true, retryAttempt + 1).ConfigureAwait(false);
            }

            throw new InvalidOperationException("Could not reach EZSheets cloud. Try again in a moment.", ex);
        }

        using (response)
        {
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

            if (response.StatusCode == HttpStatusCode.Unauthorized)
            {
                if (allowPersistedRetry && this.configuration.TryReloadCharacterSessionFromPersisted(this.loadedCharacterFullName) && this.IsAuthenticated)
                {
                    return await this.SendAuthenticatedFunctionAsync<T>(action, payload, cancellationToken, allowRefreshRetry, allowPersistedRetry: false, alreadySerialized: true, retryAttempt).ConfigureAwait(false);
                }

                if (allowRefreshRetry && !string.IsNullOrWhiteSpace(this.GetCurrentRefreshToken()))
                {
                    try
                    {
                        var refreshed = await this.RefreshDiscordSessionAsync(cancellationToken, alreadySerialized: true).ConfigureAwait(false);
                        this.StoreSession(refreshed);
                        this.suppressTokenRefreshUntilUtc = DateTimeOffset.MinValue;
                        return await this.SendAuthenticatedFunctionAsync<T>(action, payload, cancellationToken, allowRefreshRetry: false, allowPersistedRetry: false, alreadySerialized: true, retryAttempt).ConfigureAwait(false);
                    }
                    catch (SupabaseApiException ex) when (string.Equals(ex.Message, "invalid_grant", StringComparison.OrdinalIgnoreCase))
                    {
                        if (this.configuration.TryReloadCharacterSessionFromPersisted(this.loadedCharacterFullName) && this.IsAuthenticated)
                        {
                            return await this.SendAuthenticatedFunctionAsync<T>(action, payload, cancellationToken, allowRefreshRetry: false, allowPersistedRetry: false, alreadySerialized: true, retryAttempt).ConfigureAwait(false);
                        }

                        this.configuration.ClearCharacterSession(this.loadedCharacterFullName, keepRememberedLogin: false);
                        this.configuration.Save();
                        throw new InvalidOperationException("Your Discord login expired. Sign in again to continue using cloud features.", ex);
                    }
                }
            }

            if (!response.IsSuccessStatusCode)
            {
                var parsedError = ParseError(body);
                if ((response.StatusCode == HttpStatusCode.RequestTimeout || response.StatusCode == HttpStatusCode.GatewayTimeout || parsedError.Contains("statement timeout", StringComparison.OrdinalIgnoreCase))
                    && ShouldRetryAction(action, retryAttempt))
                {
                    await Task.Delay(GetRetryDelay(retryAttempt), cancellationToken).ConfigureAwait(false);
                    return await this.SendAuthenticatedFunctionAsync<T>(action, payload, cancellationToken, allowRefreshRetry, allowPersistedRetry, alreadySerialized: true, retryAttempt + 1).ConfigureAwait(false);
                }

                throw new SupabaseApiException(parsedError, response.StatusCode);
            }

            var item = JsonSerializer.Deserialize<T>(body, JsonOptions);
            if (item is null)
            {
                throw new InvalidOperationException("EZSheets returned an empty payload.");
            }

            return item;
        }
    }

    private async Task<T> SendPublicFunctionAsync<T>(string action, object? payload, CancellationToken cancellationToken, bool alreadySerialized = false, int retryAttempt = 0)
        where T : class
    {
        if (!alreadySerialized)
        {
            return await this.SendPublicFunctionAsync<T>(action, payload, cancellationToken, alreadySerialized: true, retryAttempt).ConfigureAwait(false);
        }

        using var request = this.CreateFunctionRequest(action, payload, includeAuth: false);
        using var requestCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, this.disposeCts.Token);
        requestCts.CancelAfter(GetActionTimeout(action));

        HttpResponseMessage response;
        try
        {
            response = await this.httpClient.SendAsync(request, requestCts.Token).ConfigureAwait(false);
        }
        catch (TaskCanceledException ex) when (!cancellationToken.IsCancellationRequested && !this.disposeCts.IsCancellationRequested)
        {
            if (ShouldRetryAction(action, retryAttempt))
            {
                await Task.Delay(GetRetryDelay(retryAttempt), cancellationToken).ConfigureAwait(false);
                return await this.SendPublicFunctionAsync<T>(action, payload, cancellationToken, alreadySerialized: true, retryAttempt + 1).ConfigureAwait(false);
            }

            throw new TimeoutException("EZSheets cloud request timed out. Try again in a moment.", ex);
        }
        catch (HttpRequestException ex) when (!cancellationToken.IsCancellationRequested && !this.disposeCts.IsCancellationRequested)
        {
            if (ShouldRetryAction(action, retryAttempt))
            {
                await Task.Delay(GetRetryDelay(retryAttempt), cancellationToken).ConfigureAwait(false);
                return await this.SendPublicFunctionAsync<T>(action, payload, cancellationToken, alreadySerialized: true, retryAttempt + 1).ConfigureAwait(false);
            }

            throw new InvalidOperationException("Could not reach EZSheets cloud. Try again in a moment.", ex);
        }

        using (response)
        {
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
            {
                var parsedError = ParseError(body);
                if ((response.StatusCode == HttpStatusCode.RequestTimeout || response.StatusCode == HttpStatusCode.GatewayTimeout || parsedError.Contains("statement timeout", StringComparison.OrdinalIgnoreCase))
                    && ShouldRetryAction(action, retryAttempt))
                {
                    await Task.Delay(GetRetryDelay(retryAttempt), cancellationToken).ConfigureAwait(false);
                    return await this.SendPublicFunctionAsync<T>(action, payload, cancellationToken, alreadySerialized: true, retryAttempt + 1).ConfigureAwait(false);
                }

                throw new SupabaseApiException(parsedError, response.StatusCode);
            }

            var item = JsonSerializer.Deserialize<T>(body, JsonOptions);
            if (item is null)
            {
                throw new InvalidOperationException("EZSheets returned an empty payload.");
            }

            return item;
        }
    }

    private static bool ShouldRetryAction(string action, int retryAttempt)
    {
        if (retryAttempt >= 1)
        {
            return false;
        }

        return action switch
        {
            "update_sheet" => true,
            "join_sheet_by_code" => true,
            "refresh_session" => true,
            _ => false,
        };
    }

    private static TimeSpan GetRetryDelay(int retryAttempt)
        => retryAttempt switch
        {
            <= 0 => TimeSpan.FromMilliseconds(250),
            _ => TimeSpan.FromMilliseconds(600),
        };

    private static TimeSpan GetActionTimeout(string action)
        => action switch
        {
            "update_sheet" => TimeSpan.FromSeconds(90),
            "get_sheet" => TimeSpan.FromSeconds(12),
            "list_sheets" => TimeSpan.FromSeconds(8),
            "get_access_role" => TimeSpan.FromSeconds(6),
            "get_sheet_runtime" => TimeSpan.FromSeconds(4),
            "sync_presence" => TimeSpan.FromSeconds(3),
            "acquire_cell_lock" => TimeSpan.FromSeconds(3),
            "release_cell_lock" => TimeSpan.FromSeconds(3),
            "list_sheet_members" => TimeSpan.FromSeconds(6),
            "list_sheet_blocklist" => TimeSpan.FromSeconds(6),
            "post_chat_message" => TimeSpan.FromSeconds(5),
            "refresh_session" => TimeSpan.FromSeconds(8),
            _ => TimeSpan.FromSeconds(8),
        };

    private HttpRequestMessage CreateFunctionRequest(string action, object? payload, bool includeAuth)
    {
        if (!this.HasConfiguredProject)
        {
            throw new InvalidOperationException("Supabase URL, publishable key, and Discord client id are required first.");
        }

        if (this.httpClient.BaseAddress is null)
        {
            throw new InvalidOperationException("Supabase URL is invalid.");
        }

        var request = new HttpRequestMessage(HttpMethod.Post, $"functions/v1/{Configuration.EZSheetsFunctionName}");
        request.Headers.Add("apikey", this.EffectivePublishableKey);
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        if (includeAuth)
        {
            var serverSessionToken = SanitizeHeaderValue(this.configuration.ServerSessionToken);
            var accessToken = SanitizeAccessToken(this.configuration.AccessToken);

            if (!string.IsNullOrWhiteSpace(serverSessionToken))
            {
                request.Headers.TryAddWithoutValidation("x-sheetsync-session", serverSessionToken);
                request.Headers.TryAddWithoutValidation("x-EZSheets-session", serverSessionToken);
            }

            if (!string.IsNullOrWhiteSpace(accessToken))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            }
        }

        request.Content = JsonContent.Create(new
        {
            action,
            payload,
        });

        return request;
    }

    private static string SanitizeHeaderValue(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return value.Trim().Trim('"').Replace("\r", string.Empty).Replace("\n", string.Empty);
    }

    private static string SanitizeAccessToken(string? value)
    {
        var sanitized = SanitizeHeaderValue(value);
        if (sanitized.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
        {
            sanitized = sanitized[7..].Trim();
        }

        return sanitized;
    }

    private async Task<DiscordAppSession> ExchangeDiscordCodeForSessionAsync(string code, string redirectUrl, CancellationToken cancellationToken)
    {
        return await this.SendPublicFunctionAsync<DiscordAppSession>(
            "exchange_code",
            new
            {
                code,
                redirect_uri = redirectUrl,
            },
            cancellationToken).ConfigureAwait(false);
    }

    private async Task<DiscordAppSession> RefreshDiscordSessionAsync(CancellationToken cancellationToken, bool alreadySerialized = false)
    {
        var refreshToken = this.GetCurrentRefreshToken();
        if (string.IsNullOrWhiteSpace(refreshToken))
        {
            throw new InvalidOperationException("No Discord refresh token is available.");
        }

        return await this.SendPublicFunctionAsync<DiscordAppSession>(
            "refresh_session",
            new
            {
                refresh_token = SanitizeHeaderValue(refreshToken),
                server_session_token = SanitizeHeaderValue(this.configuration.ServerSessionToken),
            },
            cancellationToken,
            alreadySerialized: alreadySerialized).ConfigureAwait(false);
    }

    private string BuildDiscordAuthorizeUrl(string redirectUrl, string state)
    {
        var builder = new UriBuilder("https://discord.com/oauth2/authorize");
        var queryParts = new List<string>
        {
            $"response_type={Uri.EscapeDataString("code")}",
            $"client_id={Uri.EscapeDataString(this.configuration.DiscordClientId.Trim())}",
            $"scope={Uri.EscapeDataString("identify")}",
            $"redirect_uri={Uri.EscapeDataString(redirectUrl)}",
            $"state={Uri.EscapeDataString(state)}",
        };
        builder.Query = string.Join("&", queryParts);
        return builder.Uri.ToString();
    }

    private void StoreSession(DiscordAppSession session)
    {
        this.EnsureCharacterSessionContextLoaded();
        this.configuration.AccessToken = SanitizeAccessToken(session.AccessToken);
        this.configuration.RefreshToken = SanitizeHeaderValue(session.RefreshToken);
        this.configuration.RememberedRefreshToken = SanitizeHeaderValue(session.RefreshToken);
        this.configuration.AccessTokenExpiresAtUnix = session.ExpiresAtUnix > 0
            ? session.ExpiresAtUnix
            : DateTimeOffset.UtcNow.AddSeconds(Math.Max(session.ExpiresIn, 0)).ToUnixTimeSeconds();
        this.configuration.ServerSessionToken = SanitizeHeaderValue(session.SessionToken);
        this.configuration.ServerSessionExpiresAtUnix = session.SessionExpiresAtUnix;
        this.configuration.UserId = session.User?.Id ?? string.Empty;
        this.configuration.UserEmail = session.User?.Email ?? string.Empty;
        this.configuration.UserDisplayName = session.User?.GetPreferredDisplayName() ?? "Discord user";
        this.PersistCurrentCharacterSession();
    }

    private void EnsureAuthenticated()
    {
        this.EnsureCharacterSessionContextLoaded();
        if (!this.IsAuthenticated)
        {
            throw new InvalidOperationException("You must be signed in before calling EZSheets online features.");
        }
    }

    private static string ParseError(string body)
    {
        try
        {
            var parsed = JsonSerializer.Deserialize<ApiErrorResponse>(body, JsonOptions);
            if (parsed is not null)
            {
                return parsed.ToFriendlyMessage();
            }
        }
        catch
        {
            // Fall back to sanitized body.
        }

        if (string.IsNullOrWhiteSpace(body))
        {
            return "EZSheets returned an unknown error.";
        }

        var trimmed = body.Trim();
        if (trimmed.StartsWith("<!DOCTYPE html", StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith("<html", StringComparison.OrdinalIgnoreCase)
            || trimmed.Contains("<title>", StringComparison.OrdinalIgnoreCase))
        {
            return "EZSheets could not reach the cloud service right now. Please try again in a moment.";
        }

        return trimmed.Length > 260 ? trimmed[..260] + "..." : trimmed;
    }

    public static string GenerateShareCode(int length = 10)
    {
        const string alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
        Span<byte> bytes = stackalloc byte[length];
        RandomNumberGenerator.Fill(bytes);
        var builder = new StringBuilder(length);
        foreach (var value in bytes)
        {
            builder.Append(alphabet[value % alphabet.Length]);
        }

        return builder.ToString();
    }

    private static string GenerateStateToken(int byteLength = 32)
    {
        Span<byte> bytes = stackalloc byte[byteLength];
        RandomNumberGenerator.Fill(bytes);
        return Base64UrlEncode(bytes.ToArray());
    }

    private static string Base64UrlEncode(byte[] bytes)
    {
        return Convert.ToBase64String(bytes)
            .TrimEnd('=')
            .Replace('+', '-')
            .Replace('/', '_');
    }

    private static async Task<LoopbackCallback> ReadLoopbackRequestAsync(TcpClient client, CancellationToken cancellationToken)
    {
        var stream = client.GetStream();
        using var memory = new MemoryStream();
        var buffer = new byte[1024];

        while (true)
        {
            var read = await stream.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken).ConfigureAwait(false);
            if (read <= 0)
            {
                throw new InvalidOperationException("EZSheets did not receive a valid OAuth callback request.");
            }

            await memory.WriteAsync(buffer.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
            if (ContainsHttpHeaderTerminator(memory.GetBuffer(), (int)memory.Length))
            {
                break;
            }

            if (memory.Length > 32 * 1024)
            {
                throw new InvalidOperationException("EZSheets received an unexpectedly large callback request.");
            }
        }

        var requestText = Encoding.ASCII.GetString(memory.GetBuffer(), 0, (int)memory.Length);
        var requestLineEnd = requestText.IndexOf("\r\n", StringComparison.Ordinal);
        var requestLine = requestLineEnd >= 0 ? requestText[..requestLineEnd] : requestText;
        var parts = requestLine.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length < 2)
        {
            throw new InvalidOperationException("EZSheets could not parse the OAuth callback request line.");
        }

        var requestTarget = parts[1];
        var uri = new Uri($"http://127.0.0.1{requestTarget}");
        var query = uri.Query.StartsWith('?') ? uri.Query[1..] : uri.Query;
        return new LoopbackCallback(uri.AbsolutePath, ParseQueryString(query));
    }

    private static bool ContainsHttpHeaderTerminator(byte[] buffer, int length)
    {
        if (length < 4)
        {
            return false;
        }

        for (var index = 3; index < length; index++)
        {
            if (buffer[index - 3] == (byte)'\r'
                && buffer[index - 2] == (byte)'\n'
                && buffer[index - 1] == (byte)'\r'
                && buffer[index] == (byte)'\n')
            {
                return true;
            }
        }

        return false;
    }

    private static async Task WriteHtmlAsync(Stream stream, string title, string message, HttpStatusCode statusCode = HttpStatusCode.OK)
    {
        var html = $"<!DOCTYPE html><html><head><meta charset=\"utf-8\"><title>{WebUtility.HtmlEncode(title)}</title></head><body style=\"font-family:Segoe UI,Arial,sans-serif;background:#0f1115;color:#f3f3f3;padding:32px\"><h2>{WebUtility.HtmlEncode(title)}</h2><p>{WebUtility.HtmlEncode(message)}</p></body></html>";
        var bodyBytes = Encoding.UTF8.GetBytes(html);
        var header = $"HTTP/1.1 {(int)statusCode} {statusCode}\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: {bodyBytes.Length}\r\nConnection: close\r\n\r\n";
        var headerBytes = Encoding.ASCII.GetBytes(header);
        await stream.WriteAsync(headerBytes).ConfigureAwait(false);
        await stream.WriteAsync(bodyBytes).ConfigureAwait(false);
        await stream.FlushAsync().ConfigureAwait(false);
    }

    private static void OpenInBrowser(string url)
    {
        var psi = new ProcessStartInfo
        {
            FileName = url,
            UseShellExecute = true,
        };

        Process.Start(psi);
    }

    private static Dictionary<string, string> ParseQueryString(string query)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(query))
        {
            return result;
        }

        var parts = query.Split('&', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var part in parts)
        {
            var separatorIndex = part.IndexOf('=');
            if (separatorIndex < 0)
            {
                result[Uri.UnescapeDataString(part)] = string.Empty;
                continue;
            }

            var key = Uri.UnescapeDataString(part[..separatorIndex]);
            var value = Uri.UnescapeDataString(part[(separatorIndex + 1)..]);
            result[key] = value;
        }

        return result;
    }

    private sealed record LoopbackCallback(string Path, Dictionary<string, string> Parameters);

    private sealed class AccessRoleResponse
    {
        [JsonPropertyName("role")]
        public string Role { get; set; } = "viewer";
    }

    private sealed class JoinSheetResponse
    {
        [JsonPropertyName("sheet_id")]
        public string SheetId { get; set; } = string.Empty;
    }

    private sealed class FunctionSuccessResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }
    }
}

public sealed class SupabaseApiException : Exception
{
    public SupabaseApiException(string message, HttpStatusCode statusCode)
        : base(message)
    {
        this.StatusCode = statusCode;
    }

    public HttpStatusCode StatusCode { get; }
}
